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
    /// Namespace prefix: a
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /// Namespace prefix: cdr
    pub const CDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
    /// Namespace prefix: dchrt
    pub const DCHRT: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    /// Namespace prefix: ddgrm
    pub const DDGRM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum LayoutTargetType {
    #[serde(rename = "inner")]
    Inner,
    #[serde(rename = "outer")]
    Outer,
}

impl std::fmt::Display for LayoutTargetType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Inner => write!(f, "inner"),
            Self::Outer => write!(f, "outer"),
        }
    }
}

impl std::str::FromStr for LayoutTargetType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "inner" => Ok(Self::Inner),
            "outer" => Ok(Self::Outer),
            _ => Err(format!("unknown LayoutTargetType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum LayoutModeType {
    #[serde(rename = "edge")]
    Edge,
    #[serde(rename = "factor")]
    Factor,
}

impl std::fmt::Display for LayoutModeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Edge => write!(f, "edge"),
            Self::Factor => write!(f, "factor"),
        }
    }
}

impl std::str::FromStr for LayoutModeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "edge" => Ok(Self::Edge),
            "factor" => Ok(Self::Factor),
            _ => Err(format!("unknown LayoutModeType value: {}", s)),
        }
    }
}

pub type RotXValue = i8;

pub type HPercentValue = String;

pub type HPercentWithSymbol = String;

pub type HPercentUShort = u16;

pub type RotYValue = u16;

pub type DepthPercentValue = String;

pub type DepthPercentWithSymbol = String;

pub type DepthPercentUShort = u16;

pub type PerspectiveValue = u8;

pub type ChartThicknessValue = String;

pub type ThicknessPercentStr = String;

pub type GapAmountValue = String;

pub type GapAmountPercent = String;

pub type GapAmountUShort = u16;

pub type OverlapValue = String;

pub type OverlapPercent = String;

pub type OverlapByte = i8;

pub type BubbleScaleValue = String;

pub type BubbleScalePercent = String;

pub type BubbleScaleUInt = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SizeRepresentsType {
    #[serde(rename = "area")]
    Area,
    #[serde(rename = "w")]
    W,
}

impl std::fmt::Display for SizeRepresentsType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Area => write!(f, "area"),
            Self::W => write!(f, "w"),
        }
    }
}

impl std::str::FromStr for SizeRepresentsType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "area" => Ok(Self::Area),
            "w" => Ok(Self::W),
            _ => Err(format!("unknown SizeRepresentsType value: {}", s)),
        }
    }
}

pub type FirstSliceAngValue = u16;

pub type HoleSizeValue = String;

pub type HoleSizePercent = String;

pub type HoleSizeUByte = u8;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SplitTypeValue {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "cust")]
    Cust,
    #[serde(rename = "percent")]
    Percent,
    #[serde(rename = "pos")]
    Pos,
    #[serde(rename = "val")]
    Val,
}

impl std::fmt::Display for SplitTypeValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::Cust => write!(f, "cust"),
            Self::Percent => write!(f, "percent"),
            Self::Pos => write!(f, "pos"),
            Self::Val => write!(f, "val"),
        }
    }
}

impl std::str::FromStr for SplitTypeValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "cust" => Ok(Self::Cust),
            "percent" => Ok(Self::Percent),
            "pos" => Ok(Self::Pos),
            "val" => Ok(Self::Val),
            _ => Err(format!("unknown SplitTypeValue value: {}", s)),
        }
    }
}

pub type SecondPieSizeValue = String;

pub type SecondPieSizePercent = String;

pub type SecondPieSizeUShort = u16;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum LabelAlignType {
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
}

impl std::fmt::Display for LabelAlignType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ctr => write!(f, "ctr"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
        }
    }
}

impl std::str::FromStr for LabelAlignType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ctr" => Ok(Self::Ctr),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            _ => Err(format!("unknown LabelAlignType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DataLabelPositionType {
    #[serde(rename = "bestFit")]
    BestFit,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "inBase")]
    InBase,
    #[serde(rename = "inEnd")]
    InEnd,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "outEnd")]
    OutEnd,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "t")]
    T,
}

impl std::fmt::Display for DataLabelPositionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::BestFit => write!(f, "bestFit"),
            Self::B => write!(f, "b"),
            Self::Ctr => write!(f, "ctr"),
            Self::InBase => write!(f, "inBase"),
            Self::InEnd => write!(f, "inEnd"),
            Self::L => write!(f, "l"),
            Self::OutEnd => write!(f, "outEnd"),
            Self::R => write!(f, "r"),
            Self::T => write!(f, "t"),
        }
    }
}

impl std::str::FromStr for DataLabelPositionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bestFit" => Ok(Self::BestFit),
            "b" => Ok(Self::B),
            "ctr" => Ok(Self::Ctr),
            "inBase" => Ok(Self::InBase),
            "inEnd" => Ok(Self::InEnd),
            "l" => Ok(Self::L),
            "outEnd" => Ok(Self::OutEnd),
            "r" => Ok(Self::R),
            "t" => Ok(Self::T),
            _ => Err(format!("unknown DataLabelPositionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum MarkerStyleType {
    #[serde(rename = "circle")]
    Circle,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "diamond")]
    Diamond,
    #[serde(rename = "dot")]
    Dot,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "picture")]
    Picture,
    #[serde(rename = "plus")]
    Plus,
    #[serde(rename = "square")]
    Square,
    #[serde(rename = "star")]
    Star,
    #[serde(rename = "triangle")]
    Triangle,
    #[serde(rename = "x")]
    X,
    #[serde(rename = "auto")]
    Auto,
}

impl std::fmt::Display for MarkerStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Circle => write!(f, "circle"),
            Self::Dash => write!(f, "dash"),
            Self::Diamond => write!(f, "diamond"),
            Self::Dot => write!(f, "dot"),
            Self::None => write!(f, "none"),
            Self::Picture => write!(f, "picture"),
            Self::Plus => write!(f, "plus"),
            Self::Square => write!(f, "square"),
            Self::Star => write!(f, "star"),
            Self::Triangle => write!(f, "triangle"),
            Self::X => write!(f, "x"),
            Self::Auto => write!(f, "auto"),
        }
    }
}

impl std::str::FromStr for MarkerStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "circle" => Ok(Self::Circle),
            "dash" => Ok(Self::Dash),
            "diamond" => Ok(Self::Diamond),
            "dot" => Ok(Self::Dot),
            "none" => Ok(Self::None),
            "picture" => Ok(Self::Picture),
            "plus" => Ok(Self::Plus),
            "square" => Ok(Self::Square),
            "star" => Ok(Self::Star),
            "triangle" => Ok(Self::Triangle),
            "x" => Ok(Self::X),
            "auto" => Ok(Self::Auto),
            _ => Err(format!("unknown MarkerStyleType value: {}", s)),
        }
    }
}

pub type MarkerSizeValue = u8;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TrendlineTypeValue {
    #[serde(rename = "exp")]
    Exp,
    #[serde(rename = "linear")]
    Linear,
    #[serde(rename = "log")]
    Log,
    #[serde(rename = "movingAvg")]
    MovingAvg,
    #[serde(rename = "poly")]
    Poly,
    #[serde(rename = "power")]
    Power,
}

impl std::fmt::Display for TrendlineTypeValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Exp => write!(f, "exp"),
            Self::Linear => write!(f, "linear"),
            Self::Log => write!(f, "log"),
            Self::MovingAvg => write!(f, "movingAvg"),
            Self::Poly => write!(f, "poly"),
            Self::Power => write!(f, "power"),
        }
    }
}

impl std::str::FromStr for TrendlineTypeValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "exp" => Ok(Self::Exp),
            "linear" => Ok(Self::Linear),
            "log" => Ok(Self::Log),
            "movingAvg" => Ok(Self::MovingAvg),
            "poly" => Ok(Self::Poly),
            "power" => Ok(Self::Power),
            _ => Err(format!("unknown TrendlineTypeValue value: {}", s)),
        }
    }
}

pub type TrendlineOrderValue = u8;

pub type TrendlinePeriodValue = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ErrorDirectionType {
    #[serde(rename = "x")]
    X,
    #[serde(rename = "y")]
    Y,
}

impl std::fmt::Display for ErrorDirectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::X => write!(f, "x"),
            Self::Y => write!(f, "y"),
        }
    }
}

impl std::str::FromStr for ErrorDirectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "x" => Ok(Self::X),
            "y" => Ok(Self::Y),
            _ => Err(format!("unknown ErrorDirectionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ErrorBarTypeValue {
    #[serde(rename = "both")]
    Both,
    #[serde(rename = "minus")]
    Minus,
    #[serde(rename = "plus")]
    Plus,
}

impl std::fmt::Display for ErrorBarTypeValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Both => write!(f, "both"),
            Self::Minus => write!(f, "minus"),
            Self::Plus => write!(f, "plus"),
        }
    }
}

impl std::str::FromStr for ErrorBarTypeValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "both" => Ok(Self::Both),
            "minus" => Ok(Self::Minus),
            "plus" => Ok(Self::Plus),
            _ => Err(format!("unknown ErrorBarTypeValue value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ErrorValueTypeValue {
    #[serde(rename = "cust")]
    Cust,
    #[serde(rename = "fixedVal")]
    FixedVal,
    #[serde(rename = "percentage")]
    Percentage,
    #[serde(rename = "stdDev")]
    StdDev,
    #[serde(rename = "stdErr")]
    StdErr,
}

impl std::fmt::Display for ErrorValueTypeValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cust => write!(f, "cust"),
            Self::FixedVal => write!(f, "fixedVal"),
            Self::Percentage => write!(f, "percentage"),
            Self::StdDev => write!(f, "stdDev"),
            Self::StdErr => write!(f, "stdErr"),
        }
    }
}

impl std::str::FromStr for ErrorValueTypeValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cust" => Ok(Self::Cust),
            "fixedVal" => Ok(Self::FixedVal),
            "percentage" => Ok(Self::Percentage),
            "stdDev" => Ok(Self::StdDev),
            "stdErr" => Ok(Self::StdErr),
            _ => Err(format!("unknown ErrorValueTypeValue value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum GroupingType {
    #[serde(rename = "percentStacked")]
    PercentStacked,
    #[serde(rename = "standard")]
    Standard,
    #[serde(rename = "stacked")]
    Stacked,
}

impl std::fmt::Display for GroupingType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::PercentStacked => write!(f, "percentStacked"),
            Self::Standard => write!(f, "standard"),
            Self::Stacked => write!(f, "stacked"),
        }
    }
}

impl std::str::FromStr for GroupingType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "percentStacked" => Ok(Self::PercentStacked),
            "standard" => Ok(Self::Standard),
            "stacked" => Ok(Self::Stacked),
            _ => Err(format!("unknown GroupingType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ScatterStyleType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "line")]
    Line,
    #[serde(rename = "lineMarker")]
    LineMarker,
    #[serde(rename = "marker")]
    Marker,
    #[serde(rename = "smooth")]
    Smooth,
    #[serde(rename = "smoothMarker")]
    SmoothMarker,
}

impl std::fmt::Display for ScatterStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Line => write!(f, "line"),
            Self::LineMarker => write!(f, "lineMarker"),
            Self::Marker => write!(f, "marker"),
            Self::Smooth => write!(f, "smooth"),
            Self::SmoothMarker => write!(f, "smoothMarker"),
        }
    }
}

impl std::str::FromStr for ScatterStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "line" => Ok(Self::Line),
            "lineMarker" => Ok(Self::LineMarker),
            "marker" => Ok(Self::Marker),
            "smooth" => Ok(Self::Smooth),
            "smoothMarker" => Ok(Self::SmoothMarker),
            _ => Err(format!("unknown ScatterStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RadarStyleType {
    #[serde(rename = "standard")]
    Standard,
    #[serde(rename = "marker")]
    Marker,
    #[serde(rename = "filled")]
    Filled,
}

impl std::fmt::Display for RadarStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Standard => write!(f, "standard"),
            Self::Marker => write!(f, "marker"),
            Self::Filled => write!(f, "filled"),
        }
    }
}

impl std::str::FromStr for RadarStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "standard" => Ok(Self::Standard),
            "marker" => Ok(Self::Marker),
            "filled" => Ok(Self::Filled),
            _ => Err(format!("unknown RadarStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BarGroupingType {
    #[serde(rename = "percentStacked")]
    PercentStacked,
    #[serde(rename = "clustered")]
    Clustered,
    #[serde(rename = "standard")]
    Standard,
    #[serde(rename = "stacked")]
    Stacked,
}

impl std::fmt::Display for BarGroupingType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::PercentStacked => write!(f, "percentStacked"),
            Self::Clustered => write!(f, "clustered"),
            Self::Standard => write!(f, "standard"),
            Self::Stacked => write!(f, "stacked"),
        }
    }
}

impl std::str::FromStr for BarGroupingType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "percentStacked" => Ok(Self::PercentStacked),
            "clustered" => Ok(Self::Clustered),
            "standard" => Ok(Self::Standard),
            "stacked" => Ok(Self::Stacked),
            _ => Err(format!("unknown BarGroupingType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BarDirectionType {
    #[serde(rename = "bar")]
    Bar,
    #[serde(rename = "col")]
    Col,
}

impl std::fmt::Display for BarDirectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Bar => write!(f, "bar"),
            Self::Col => write!(f, "col"),
        }
    }
}

impl std::str::FromStr for BarDirectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bar" => Ok(Self::Bar),
            "col" => Ok(Self::Col),
            _ => Err(format!("unknown BarDirectionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BarShapeType {
    #[serde(rename = "cone")]
    Cone,
    #[serde(rename = "coneToMax")]
    ConeToMax,
    #[serde(rename = "box")]
    Box,
    #[serde(rename = "cylinder")]
    Cylinder,
    #[serde(rename = "pyramid")]
    Pyramid,
    #[serde(rename = "pyramidToMax")]
    PyramidToMax,
}

impl std::fmt::Display for BarShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cone => write!(f, "cone"),
            Self::ConeToMax => write!(f, "coneToMax"),
            Self::Box => write!(f, "box"),
            Self::Cylinder => write!(f, "cylinder"),
            Self::Pyramid => write!(f, "pyramid"),
            Self::PyramidToMax => write!(f, "pyramidToMax"),
        }
    }
}

impl std::str::FromStr for BarShapeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cone" => Ok(Self::Cone),
            "coneToMax" => Ok(Self::ConeToMax),
            "box" => Ok(Self::Box),
            "cylinder" => Ok(Self::Cylinder),
            "pyramid" => Ok(Self::Pyramid),
            "pyramidToMax" => Ok(Self::PyramidToMax),
            _ => Err(format!("unknown BarShapeType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum OfPieTypeValue {
    #[serde(rename = "pie")]
    Pie,
    #[serde(rename = "bar")]
    Bar,
}

impl std::fmt::Display for OfPieTypeValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Pie => write!(f, "pie"),
            Self::Bar => write!(f, "bar"),
        }
    }
}

impl std::str::FromStr for OfPieTypeValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "pie" => Ok(Self::Pie),
            "bar" => Ok(Self::Bar),
            _ => Err(format!("unknown OfPieTypeValue value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum AxisPositionType {
    #[serde(rename = "b")]
    B,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "t")]
    T,
}

impl std::fmt::Display for AxisPositionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
            Self::T => write!(f, "t"),
        }
    }
}

impl std::str::FromStr for AxisPositionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            "t" => Ok(Self::T),
            _ => Err(format!("unknown AxisPositionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum AxisCrossesType {
    #[serde(rename = "autoZero")]
    AutoZero,
    #[serde(rename = "max")]
    Max,
    #[serde(rename = "min")]
    Min,
}

impl std::fmt::Display for AxisCrossesType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AutoZero => write!(f, "autoZero"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
        }
    }
}

impl std::str::FromStr for AxisCrossesType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "autoZero" => Ok(Self::AutoZero),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            _ => Err(format!("unknown AxisCrossesType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CrossBetweenType {
    #[serde(rename = "between")]
    Between,
    #[serde(rename = "midCat")]
    MidCat,
}

impl std::fmt::Display for CrossBetweenType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Between => write!(f, "between"),
            Self::MidCat => write!(f, "midCat"),
        }
    }
}

impl std::str::FromStr for CrossBetweenType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "between" => Ok(Self::Between),
            "midCat" => Ok(Self::MidCat),
            _ => Err(format!("unknown CrossBetweenType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TickMarkType {
    #[serde(rename = "cross")]
    Cross,
    #[serde(rename = "in")]
    In,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "out")]
    Out,
}

impl std::fmt::Display for TickMarkType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cross => write!(f, "cross"),
            Self::In => write!(f, "in"),
            Self::None => write!(f, "none"),
            Self::Out => write!(f, "out"),
        }
    }
}

impl std::str::FromStr for TickMarkType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cross" => Ok(Self::Cross),
            "in" => Ok(Self::In),
            "none" => Ok(Self::None),
            "out" => Ok(Self::Out),
            _ => Err(format!("unknown TickMarkType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TickLabelPositionType {
    #[serde(rename = "high")]
    High,
    #[serde(rename = "low")]
    Low,
    #[serde(rename = "nextTo")]
    NextTo,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for TickLabelPositionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::High => write!(f, "high"),
            Self::Low => write!(f, "low"),
            Self::NextTo => write!(f, "nextTo"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for TickLabelPositionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "high" => Ok(Self::High),
            "low" => Ok(Self::Low),
            "nextTo" => Ok(Self::NextTo),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown TickLabelPositionType value: {}", s)),
        }
    }
}

pub type AxisSkipValue = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TimeUnitType {
    #[serde(rename = "days")]
    Days,
    #[serde(rename = "months")]
    Months,
    #[serde(rename = "years")]
    Years,
}

impl std::fmt::Display for TimeUnitType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Days => write!(f, "days"),
            Self::Months => write!(f, "months"),
            Self::Years => write!(f, "years"),
        }
    }
}

impl std::str::FromStr for TimeUnitType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "days" => Ok(Self::Days),
            "months" => Ok(Self::Months),
            "years" => Ok(Self::Years),
            _ => Err(format!("unknown TimeUnitType value: {}", s)),
        }
    }
}

pub type AxisUnitValue = f64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BuiltInUnitType {
    #[serde(rename = "hundreds")]
    Hundreds,
    #[serde(rename = "thousands")]
    Thousands,
    #[serde(rename = "tenThousands")]
    TenThousands,
    #[serde(rename = "hundredThousands")]
    HundredThousands,
    #[serde(rename = "millions")]
    Millions,
    #[serde(rename = "tenMillions")]
    TenMillions,
    #[serde(rename = "hundredMillions")]
    HundredMillions,
    #[serde(rename = "billions")]
    Billions,
    #[serde(rename = "trillions")]
    Trillions,
}

impl std::fmt::Display for BuiltInUnitType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Hundreds => write!(f, "hundreds"),
            Self::Thousands => write!(f, "thousands"),
            Self::TenThousands => write!(f, "tenThousands"),
            Self::HundredThousands => write!(f, "hundredThousands"),
            Self::Millions => write!(f, "millions"),
            Self::TenMillions => write!(f, "tenMillions"),
            Self::HundredMillions => write!(f, "hundredMillions"),
            Self::Billions => write!(f, "billions"),
            Self::Trillions => write!(f, "trillions"),
        }
    }
}

impl std::str::FromStr for BuiltInUnitType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "hundreds" => Ok(Self::Hundreds),
            "thousands" => Ok(Self::Thousands),
            "tenThousands" => Ok(Self::TenThousands),
            "hundredThousands" => Ok(Self::HundredThousands),
            "millions" => Ok(Self::Millions),
            "tenMillions" => Ok(Self::TenMillions),
            "hundredMillions" => Ok(Self::HundredMillions),
            "billions" => Ok(Self::Billions),
            "trillions" => Ok(Self::Trillions),
            _ => Err(format!("unknown BuiltInUnitType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PictureFormatType {
    #[serde(rename = "stretch")]
    Stretch,
    #[serde(rename = "stack")]
    Stack,
    #[serde(rename = "stackScale")]
    StackScale,
}

impl std::fmt::Display for PictureFormatType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stretch => write!(f, "stretch"),
            Self::Stack => write!(f, "stack"),
            Self::StackScale => write!(f, "stackScale"),
        }
    }
}

impl std::str::FromStr for PictureFormatType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stretch" => Ok(Self::Stretch),
            "stack" => Ok(Self::Stack),
            "stackScale" => Ok(Self::StackScale),
            _ => Err(format!("unknown PictureFormatType value: {}", s)),
        }
    }
}

pub type PictureStackUnitValue = f64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum AxisOrientationType {
    #[serde(rename = "maxMin")]
    MaxMin,
    #[serde(rename = "minMax")]
    MinMax,
}

impl std::fmt::Display for AxisOrientationType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MaxMin => write!(f, "maxMin"),
            Self::MinMax => write!(f, "minMax"),
        }
    }
}

impl std::str::FromStr for AxisOrientationType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "maxMin" => Ok(Self::MaxMin),
            "minMax" => Ok(Self::MinMax),
            _ => Err(format!("unknown AxisOrientationType value: {}", s)),
        }
    }
}

pub type LogBaseValue = f64;

pub type LabelOffsetValue = String;

pub type LabelOffsetPercent = String;

pub type LabelOffsetUShort = u16;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum LegendPositionType {
    #[serde(rename = "b")]
    B,
    #[serde(rename = "tr")]
    Tr,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "t")]
    T,
}

impl std::fmt::Display for LegendPositionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::Tr => write!(f, "tr"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
            Self::T => write!(f, "t"),
        }
    }
}

impl std::str::FromStr for LegendPositionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "tr" => Ok(Self::Tr),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            "t" => Ok(Self::T),
            _ => Err(format!("unknown LegendPositionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DisplayBlanksAsType {
    #[serde(rename = "span")]
    Span,
    #[serde(rename = "gap")]
    Gap,
    #[serde(rename = "zero")]
    Zero,
}

impl std::fmt::Display for DisplayBlanksAsType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Span => write!(f, "span"),
            Self::Gap => write!(f, "gap"),
            Self::Zero => write!(f, "zero"),
        }
    }
}

impl std::str::FromStr for DisplayBlanksAsType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "span" => Ok(Self::Span),
            "gap" => Ok(Self::Gap),
            "zero" => Ok(Self::Zero),
            _ => Err(format!("unknown DisplayBlanksAsType value: {}", s)),
        }
    }
}

pub type ChartStyleValue = u8;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ChartPageOrientation {
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "portrait")]
    Portrait,
    #[serde(rename = "landscape")]
    Landscape,
}

impl std::fmt::Display for ChartPageOrientation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Default => write!(f, "default"),
            Self::Portrait => write!(f, "portrait"),
            Self::Landscape => write!(f, "landscape"),
        }
    }
}

impl std::str::FromStr for ChartPageOrientation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "default" => Ok(Self::Default),
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(format!("unknown ChartPageOrientation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STClrAppMethod {
    #[serde(rename = "span")]
    Span,
    #[serde(rename = "cycle")]
    Cycle,
    #[serde(rename = "repeat")]
    Repeat,
}

impl std::fmt::Display for STClrAppMethod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Span => write!(f, "span"),
            Self::Cycle => write!(f, "cycle"),
            Self::Repeat => write!(f, "repeat"),
        }
    }
}

impl std::str::FromStr for STClrAppMethod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "span" => Ok(Self::Span),
            "cycle" => Ok(Self::Cycle),
            "repeat" => Ok(Self::Repeat),
            _ => Err(format!("unknown STClrAppMethod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHueDir {
    #[serde(rename = "cw")]
    Cw,
    #[serde(rename = "ccw")]
    Ccw,
}

impl std::fmt::Display for STHueDir {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cw => write!(f, "cw"),
            Self::Ccw => write!(f, "ccw"),
        }
    }
}

impl std::str::FromStr for STHueDir {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cw" => Ok(Self::Cw),
            "ccw" => Ok(Self::Ccw),
            _ => Err(format!("unknown STHueDir value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPtType {
    #[serde(rename = "node")]
    Node,
    #[serde(rename = "asst")]
    Asst,
    #[serde(rename = "doc")]
    Doc,
    #[serde(rename = "pres")]
    Pres,
    #[serde(rename = "parTrans")]
    ParTrans,
    #[serde(rename = "sibTrans")]
    SibTrans,
}

impl std::fmt::Display for STPtType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Node => write!(f, "node"),
            Self::Asst => write!(f, "asst"),
            Self::Doc => write!(f, "doc"),
            Self::Pres => write!(f, "pres"),
            Self::ParTrans => write!(f, "parTrans"),
            Self::SibTrans => write!(f, "sibTrans"),
        }
    }
}

impl std::str::FromStr for STPtType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "node" => Ok(Self::Node),
            "asst" => Ok(Self::Asst),
            "doc" => Ok(Self::Doc),
            "pres" => Ok(Self::Pres),
            "parTrans" => Ok(Self::ParTrans),
            "sibTrans" => Ok(Self::SibTrans),
            _ => Err(format!("unknown STPtType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCxnType {
    #[serde(rename = "parOf")]
    ParOf,
    #[serde(rename = "presOf")]
    PresOf,
    #[serde(rename = "presParOf")]
    PresParOf,
    #[serde(rename = "unknownRelationship")]
    UnknownRelationship,
}

impl std::fmt::Display for STCxnType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ParOf => write!(f, "parOf"),
            Self::PresOf => write!(f, "presOf"),
            Self::PresParOf => write!(f, "presParOf"),
            Self::UnknownRelationship => write!(f, "unknownRelationship"),
        }
    }
}

impl std::str::FromStr for STCxnType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "parOf" => Ok(Self::ParOf),
            "presOf" => Ok(Self::PresOf),
            "presParOf" => Ok(Self::PresParOf),
            "unknownRelationship" => Ok(Self::UnknownRelationship),
            _ => Err(format!("unknown STCxnType value: {}", s)),
        }
    }
}

pub type STLayoutShapeType = String;

pub type STIndex1 = u32;

pub type STParameterVal = String;

pub type STModelId = String;

pub type STPrSetCustVal = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDirection {
    #[serde(rename = "norm")]
    Norm,
    #[serde(rename = "rev")]
    Rev,
}

impl std::fmt::Display for STDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Norm => write!(f, "norm"),
            Self::Rev => write!(f, "rev"),
        }
    }
}

impl std::str::FromStr for STDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "norm" => Ok(Self::Norm),
            "rev" => Ok(Self::Rev),
            _ => Err(format!("unknown STDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHierBranchStyle {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "hang")]
    Hang,
    #[serde(rename = "std")]
    Std,
    #[serde(rename = "init")]
    Init,
}

impl std::fmt::Display for STHierBranchStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
            Self::Hang => write!(f, "hang"),
            Self::Std => write!(f, "std"),
            Self::Init => write!(f, "init"),
        }
    }
}

impl std::str::FromStr for STHierBranchStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            "hang" => Ok(Self::Hang),
            "std" => Ok(Self::Std),
            "init" => Ok(Self::Init),
            _ => Err(format!("unknown STHierBranchStyle value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAnimOneStr {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "one")]
    One,
    #[serde(rename = "branch")]
    Branch,
}

impl std::fmt::Display for STAnimOneStr {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::One => write!(f, "one"),
            Self::Branch => write!(f, "branch"),
        }
    }
}

impl std::str::FromStr for STAnimOneStr {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "one" => Ok(Self::One),
            "branch" => Ok(Self::Branch),
            _ => Err(format!("unknown STAnimOneStr value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAnimLvlStr {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "lvl")]
    Lvl,
    #[serde(rename = "ctr")]
    Ctr,
}

impl std::fmt::Display for STAnimLvlStr {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Lvl => write!(f, "lvl"),
            Self::Ctr => write!(f, "ctr"),
        }
    }
}

impl std::str::FromStr for STAnimLvlStr {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "lvl" => Ok(Self::Lvl),
            "ctr" => Ok(Self::Ctr),
            _ => Err(format!("unknown STAnimLvlStr value: {}", s)),
        }
    }
}

pub type STNodeCount = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STResizeHandlesStr {
    #[serde(rename = "exact")]
    Exact,
    #[serde(rename = "rel")]
    Rel,
}

impl std::fmt::Display for STResizeHandlesStr {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Exact => write!(f, "exact"),
            Self::Rel => write!(f, "rel"),
        }
    }
}

impl std::str::FromStr for STResizeHandlesStr {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "exact" => Ok(Self::Exact),
            "rel" => Ok(Self::Rel),
            _ => Err(format!("unknown STResizeHandlesStr value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAlgorithmType {
    #[serde(rename = "composite")]
    Composite,
    #[serde(rename = "conn")]
    Conn,
    #[serde(rename = "cycle")]
    Cycle,
    #[serde(rename = "hierChild")]
    HierChild,
    #[serde(rename = "hierRoot")]
    HierRoot,
    #[serde(rename = "pyra")]
    Pyra,
    #[serde(rename = "lin")]
    Lin,
    #[serde(rename = "sp")]
    Sp,
    #[serde(rename = "tx")]
    Tx,
    #[serde(rename = "snake")]
    Snake,
}

impl std::fmt::Display for STAlgorithmType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Composite => write!(f, "composite"),
            Self::Conn => write!(f, "conn"),
            Self::Cycle => write!(f, "cycle"),
            Self::HierChild => write!(f, "hierChild"),
            Self::HierRoot => write!(f, "hierRoot"),
            Self::Pyra => write!(f, "pyra"),
            Self::Lin => write!(f, "lin"),
            Self::Sp => write!(f, "sp"),
            Self::Tx => write!(f, "tx"),
            Self::Snake => write!(f, "snake"),
        }
    }
}

impl std::str::FromStr for STAlgorithmType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "composite" => Ok(Self::Composite),
            "conn" => Ok(Self::Conn),
            "cycle" => Ok(Self::Cycle),
            "hierChild" => Ok(Self::HierChild),
            "hierRoot" => Ok(Self::HierRoot),
            "pyra" => Ok(Self::Pyra),
            "lin" => Ok(Self::Lin),
            "sp" => Ok(Self::Sp),
            "tx" => Ok(Self::Tx),
            "snake" => Ok(Self::Snake),
            _ => Err(format!("unknown STAlgorithmType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAxisType {
    #[serde(rename = "self")]
    SelfNode,
    #[serde(rename = "ch")]
    Ch,
    #[serde(rename = "des")]
    Des,
    #[serde(rename = "desOrSelf")]
    DesOrSelf,
    #[serde(rename = "par")]
    Par,
    #[serde(rename = "ancst")]
    Ancst,
    #[serde(rename = "ancstOrSelf")]
    AncstOrSelf,
    #[serde(rename = "followSib")]
    FollowSib,
    #[serde(rename = "precedSib")]
    PrecedSib,
    #[serde(rename = "follow")]
    Follow,
    #[serde(rename = "preced")]
    Preced,
    #[serde(rename = "root")]
    Root,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STAxisType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SelfNode => write!(f, "self"),
            Self::Ch => write!(f, "ch"),
            Self::Des => write!(f, "des"),
            Self::DesOrSelf => write!(f, "desOrSelf"),
            Self::Par => write!(f, "par"),
            Self::Ancst => write!(f, "ancst"),
            Self::AncstOrSelf => write!(f, "ancstOrSelf"),
            Self::FollowSib => write!(f, "followSib"),
            Self::PrecedSib => write!(f, "precedSib"),
            Self::Follow => write!(f, "follow"),
            Self::Preced => write!(f, "preced"),
            Self::Root => write!(f, "root"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STAxisType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "self" => Ok(Self::SelfNode),
            "ch" => Ok(Self::Ch),
            "des" => Ok(Self::Des),
            "desOrSelf" => Ok(Self::DesOrSelf),
            "par" => Ok(Self::Par),
            "ancst" => Ok(Self::Ancst),
            "ancstOrSelf" => Ok(Self::AncstOrSelf),
            "followSib" => Ok(Self::FollowSib),
            "precedSib" => Ok(Self::PrecedSib),
            "follow" => Ok(Self::Follow),
            "preced" => Ok(Self::Preced),
            "root" => Ok(Self::Root),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STAxisType value: {}", s)),
        }
    }
}

pub type STAxisTypes = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBoolOperator {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "equ")]
    Equ,
    #[serde(rename = "gte")]
    Gte,
    #[serde(rename = "lte")]
    Lte,
}

impl std::fmt::Display for STBoolOperator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Equ => write!(f, "equ"),
            Self::Gte => write!(f, "gte"),
            Self::Lte => write!(f, "lte"),
        }
    }
}

impl std::str::FromStr for STBoolOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "equ" => Ok(Self::Equ),
            "gte" => Ok(Self::Gte),
            "lte" => Ok(Self::Lte),
            _ => Err(format!("unknown STBoolOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STChildOrderType {
    #[serde(rename = "b")]
    B,
    #[serde(rename = "t")]
    T,
}

impl std::fmt::Display for STChildOrderType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::T => write!(f, "t"),
        }
    }
}

impl std::str::FromStr for STChildOrderType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "t" => Ok(Self::T),
            _ => Err(format!("unknown STChildOrderType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConstraintType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "alignOff")]
    AlignOff,
    #[serde(rename = "begMarg")]
    BegMarg,
    #[serde(rename = "bendDist")]
    BendDist,
    #[serde(rename = "begPad")]
    BegPad,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "bMarg")]
    BMarg,
    #[serde(rename = "bOff")]
    BOff,
    #[serde(rename = "ctrX")]
    CtrX,
    #[serde(rename = "ctrXOff")]
    CtrXOff,
    #[serde(rename = "ctrY")]
    CtrY,
    #[serde(rename = "ctrYOff")]
    CtrYOff,
    #[serde(rename = "connDist")]
    ConnDist,
    #[serde(rename = "diam")]
    Diam,
    #[serde(rename = "endMarg")]
    EndMarg,
    #[serde(rename = "endPad")]
    EndPad,
    #[serde(rename = "h")]
    H,
    #[serde(rename = "hArH")]
    HArH,
    #[serde(rename = "hOff")]
    HOff,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "lMarg")]
    LMarg,
    #[serde(rename = "lOff")]
    LOff,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "rMarg")]
    RMarg,
    #[serde(rename = "rOff")]
    ROff,
    #[serde(rename = "primFontSz")]
    PrimFontSz,
    #[serde(rename = "pyraAcctRatio")]
    PyraAcctRatio,
    #[serde(rename = "secFontSz")]
    SecFontSz,
    #[serde(rename = "sibSp")]
    SibSp,
    #[serde(rename = "secSibSp")]
    SecSibSp,
    #[serde(rename = "sp")]
    Sp,
    #[serde(rename = "stemThick")]
    StemThick,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "tMarg")]
    TMarg,
    #[serde(rename = "tOff")]
    TOff,
    #[serde(rename = "userA")]
    UserA,
    #[serde(rename = "userB")]
    UserB,
    #[serde(rename = "userC")]
    UserC,
    #[serde(rename = "userD")]
    UserD,
    #[serde(rename = "userE")]
    UserE,
    #[serde(rename = "userF")]
    UserF,
    #[serde(rename = "userG")]
    UserG,
    #[serde(rename = "userH")]
    UserH,
    #[serde(rename = "userI")]
    UserI,
    #[serde(rename = "userJ")]
    UserJ,
    #[serde(rename = "userK")]
    UserK,
    #[serde(rename = "userL")]
    UserL,
    #[serde(rename = "userM")]
    UserM,
    #[serde(rename = "userN")]
    UserN,
    #[serde(rename = "userO")]
    UserO,
    #[serde(rename = "userP")]
    UserP,
    #[serde(rename = "userQ")]
    UserQ,
    #[serde(rename = "userR")]
    UserR,
    #[serde(rename = "userS")]
    UserS,
    #[serde(rename = "userT")]
    UserT,
    #[serde(rename = "userU")]
    UserU,
    #[serde(rename = "userV")]
    UserV,
    #[serde(rename = "userW")]
    UserW,
    #[serde(rename = "userX")]
    UserX,
    #[serde(rename = "userY")]
    UserY,
    #[serde(rename = "userZ")]
    UserZ,
    #[serde(rename = "w")]
    W,
    #[serde(rename = "wArH")]
    WArH,
    #[serde(rename = "wOff")]
    WOff,
}

impl std::fmt::Display for STConstraintType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::AlignOff => write!(f, "alignOff"),
            Self::BegMarg => write!(f, "begMarg"),
            Self::BendDist => write!(f, "bendDist"),
            Self::BegPad => write!(f, "begPad"),
            Self::B => write!(f, "b"),
            Self::BMarg => write!(f, "bMarg"),
            Self::BOff => write!(f, "bOff"),
            Self::CtrX => write!(f, "ctrX"),
            Self::CtrXOff => write!(f, "ctrXOff"),
            Self::CtrY => write!(f, "ctrY"),
            Self::CtrYOff => write!(f, "ctrYOff"),
            Self::ConnDist => write!(f, "connDist"),
            Self::Diam => write!(f, "diam"),
            Self::EndMarg => write!(f, "endMarg"),
            Self::EndPad => write!(f, "endPad"),
            Self::H => write!(f, "h"),
            Self::HArH => write!(f, "hArH"),
            Self::HOff => write!(f, "hOff"),
            Self::L => write!(f, "l"),
            Self::LMarg => write!(f, "lMarg"),
            Self::LOff => write!(f, "lOff"),
            Self::R => write!(f, "r"),
            Self::RMarg => write!(f, "rMarg"),
            Self::ROff => write!(f, "rOff"),
            Self::PrimFontSz => write!(f, "primFontSz"),
            Self::PyraAcctRatio => write!(f, "pyraAcctRatio"),
            Self::SecFontSz => write!(f, "secFontSz"),
            Self::SibSp => write!(f, "sibSp"),
            Self::SecSibSp => write!(f, "secSibSp"),
            Self::Sp => write!(f, "sp"),
            Self::StemThick => write!(f, "stemThick"),
            Self::T => write!(f, "t"),
            Self::TMarg => write!(f, "tMarg"),
            Self::TOff => write!(f, "tOff"),
            Self::UserA => write!(f, "userA"),
            Self::UserB => write!(f, "userB"),
            Self::UserC => write!(f, "userC"),
            Self::UserD => write!(f, "userD"),
            Self::UserE => write!(f, "userE"),
            Self::UserF => write!(f, "userF"),
            Self::UserG => write!(f, "userG"),
            Self::UserH => write!(f, "userH"),
            Self::UserI => write!(f, "userI"),
            Self::UserJ => write!(f, "userJ"),
            Self::UserK => write!(f, "userK"),
            Self::UserL => write!(f, "userL"),
            Self::UserM => write!(f, "userM"),
            Self::UserN => write!(f, "userN"),
            Self::UserO => write!(f, "userO"),
            Self::UserP => write!(f, "userP"),
            Self::UserQ => write!(f, "userQ"),
            Self::UserR => write!(f, "userR"),
            Self::UserS => write!(f, "userS"),
            Self::UserT => write!(f, "userT"),
            Self::UserU => write!(f, "userU"),
            Self::UserV => write!(f, "userV"),
            Self::UserW => write!(f, "userW"),
            Self::UserX => write!(f, "userX"),
            Self::UserY => write!(f, "userY"),
            Self::UserZ => write!(f, "userZ"),
            Self::W => write!(f, "w"),
            Self::WArH => write!(f, "wArH"),
            Self::WOff => write!(f, "wOff"),
        }
    }
}

impl std::str::FromStr for STConstraintType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "alignOff" => Ok(Self::AlignOff),
            "begMarg" => Ok(Self::BegMarg),
            "bendDist" => Ok(Self::BendDist),
            "begPad" => Ok(Self::BegPad),
            "b" => Ok(Self::B),
            "bMarg" => Ok(Self::BMarg),
            "bOff" => Ok(Self::BOff),
            "ctrX" => Ok(Self::CtrX),
            "ctrXOff" => Ok(Self::CtrXOff),
            "ctrY" => Ok(Self::CtrY),
            "ctrYOff" => Ok(Self::CtrYOff),
            "connDist" => Ok(Self::ConnDist),
            "diam" => Ok(Self::Diam),
            "endMarg" => Ok(Self::EndMarg),
            "endPad" => Ok(Self::EndPad),
            "h" => Ok(Self::H),
            "hArH" => Ok(Self::HArH),
            "hOff" => Ok(Self::HOff),
            "l" => Ok(Self::L),
            "lMarg" => Ok(Self::LMarg),
            "lOff" => Ok(Self::LOff),
            "r" => Ok(Self::R),
            "rMarg" => Ok(Self::RMarg),
            "rOff" => Ok(Self::ROff),
            "primFontSz" => Ok(Self::PrimFontSz),
            "pyraAcctRatio" => Ok(Self::PyraAcctRatio),
            "secFontSz" => Ok(Self::SecFontSz),
            "sibSp" => Ok(Self::SibSp),
            "secSibSp" => Ok(Self::SecSibSp),
            "sp" => Ok(Self::Sp),
            "stemThick" => Ok(Self::StemThick),
            "t" => Ok(Self::T),
            "tMarg" => Ok(Self::TMarg),
            "tOff" => Ok(Self::TOff),
            "userA" => Ok(Self::UserA),
            "userB" => Ok(Self::UserB),
            "userC" => Ok(Self::UserC),
            "userD" => Ok(Self::UserD),
            "userE" => Ok(Self::UserE),
            "userF" => Ok(Self::UserF),
            "userG" => Ok(Self::UserG),
            "userH" => Ok(Self::UserH),
            "userI" => Ok(Self::UserI),
            "userJ" => Ok(Self::UserJ),
            "userK" => Ok(Self::UserK),
            "userL" => Ok(Self::UserL),
            "userM" => Ok(Self::UserM),
            "userN" => Ok(Self::UserN),
            "userO" => Ok(Self::UserO),
            "userP" => Ok(Self::UserP),
            "userQ" => Ok(Self::UserQ),
            "userR" => Ok(Self::UserR),
            "userS" => Ok(Self::UserS),
            "userT" => Ok(Self::UserT),
            "userU" => Ok(Self::UserU),
            "userV" => Ok(Self::UserV),
            "userW" => Ok(Self::UserW),
            "userX" => Ok(Self::UserX),
            "userY" => Ok(Self::UserY),
            "userZ" => Ok(Self::UserZ),
            "w" => Ok(Self::W),
            "wArH" => Ok(Self::WArH),
            "wOff" => Ok(Self::WOff),
            _ => Err(format!("unknown STConstraintType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConstraintRelationship {
    #[serde(rename = "self")]
    SelfNode,
    #[serde(rename = "ch")]
    Ch,
    #[serde(rename = "des")]
    Des,
}

impl std::fmt::Display for STConstraintRelationship {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SelfNode => write!(f, "self"),
            Self::Ch => write!(f, "ch"),
            Self::Des => write!(f, "des"),
        }
    }
}

impl std::str::FromStr for STConstraintRelationship {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "self" => Ok(Self::SelfNode),
            "ch" => Ok(Self::Ch),
            "des" => Ok(Self::Des),
            _ => Err(format!("unknown STConstraintRelationship value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STElementType {
    #[serde(rename = "all")]
    All,
    #[serde(rename = "doc")]
    Doc,
    #[serde(rename = "node")]
    Node,
    #[serde(rename = "norm")]
    Norm,
    #[serde(rename = "nonNorm")]
    NonNorm,
    #[serde(rename = "asst")]
    Asst,
    #[serde(rename = "nonAsst")]
    NonAsst,
    #[serde(rename = "parTrans")]
    ParTrans,
    #[serde(rename = "pres")]
    Pres,
    #[serde(rename = "sibTrans")]
    SibTrans,
}

impl std::fmt::Display for STElementType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::All => write!(f, "all"),
            Self::Doc => write!(f, "doc"),
            Self::Node => write!(f, "node"),
            Self::Norm => write!(f, "norm"),
            Self::NonNorm => write!(f, "nonNorm"),
            Self::Asst => write!(f, "asst"),
            Self::NonAsst => write!(f, "nonAsst"),
            Self::ParTrans => write!(f, "parTrans"),
            Self::Pres => write!(f, "pres"),
            Self::SibTrans => write!(f, "sibTrans"),
        }
    }
}

impl std::str::FromStr for STElementType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "all" => Ok(Self::All),
            "doc" => Ok(Self::Doc),
            "node" => Ok(Self::Node),
            "norm" => Ok(Self::Norm),
            "nonNorm" => Ok(Self::NonNorm),
            "asst" => Ok(Self::Asst),
            "nonAsst" => Ok(Self::NonAsst),
            "parTrans" => Ok(Self::ParTrans),
            "pres" => Ok(Self::Pres),
            "sibTrans" => Ok(Self::SibTrans),
            _ => Err(format!("unknown STElementType value: {}", s)),
        }
    }
}

pub type STElementTypes = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STParameterId {
    #[serde(rename = "horzAlign")]
    HorzAlign,
    #[serde(rename = "vertAlign")]
    VertAlign,
    #[serde(rename = "chDir")]
    ChDir,
    #[serde(rename = "chAlign")]
    ChAlign,
    #[serde(rename = "secChAlign")]
    SecChAlign,
    #[serde(rename = "linDir")]
    LinDir,
    #[serde(rename = "secLinDir")]
    SecLinDir,
    #[serde(rename = "stElem")]
    StElem,
    #[serde(rename = "bendPt")]
    BendPt,
    #[serde(rename = "connRout")]
    ConnRout,
    #[serde(rename = "begSty")]
    BegSty,
    #[serde(rename = "endSty")]
    EndSty,
    #[serde(rename = "dim")]
    Dim,
    #[serde(rename = "rotPath")]
    RotPath,
    #[serde(rename = "ctrShpMap")]
    CtrShpMap,
    #[serde(rename = "nodeHorzAlign")]
    NodeHorzAlign,
    #[serde(rename = "nodeVertAlign")]
    NodeVertAlign,
    #[serde(rename = "fallback")]
    Fallback,
    #[serde(rename = "txDir")]
    TxDir,
    #[serde(rename = "pyraAcctPos")]
    PyraAcctPos,
    #[serde(rename = "pyraAcctTxMar")]
    PyraAcctTxMar,
    #[serde(rename = "txBlDir")]
    TxBlDir,
    #[serde(rename = "txAnchorHorz")]
    TxAnchorHorz,
    #[serde(rename = "txAnchorVert")]
    TxAnchorVert,
    #[serde(rename = "txAnchorHorzCh")]
    TxAnchorHorzCh,
    #[serde(rename = "txAnchorVertCh")]
    TxAnchorVertCh,
    #[serde(rename = "parTxLTRAlign")]
    ParTxLTRAlign,
    #[serde(rename = "parTxRTLAlign")]
    ParTxRTLAlign,
    #[serde(rename = "shpTxLTRAlignCh")]
    ShpTxLTRAlignCh,
    #[serde(rename = "shpTxRTLAlignCh")]
    ShpTxRTLAlignCh,
    #[serde(rename = "autoTxRot")]
    AutoTxRot,
    #[serde(rename = "grDir")]
    GrDir,
    #[serde(rename = "flowDir")]
    FlowDir,
    #[serde(rename = "contDir")]
    ContDir,
    #[serde(rename = "bkpt")]
    Bkpt,
    #[serde(rename = "off")]
    Off,
    #[serde(rename = "hierAlign")]
    HierAlign,
    #[serde(rename = "bkPtFixedVal")]
    BkPtFixedVal,
    #[serde(rename = "stBulletLvl")]
    StBulletLvl,
    #[serde(rename = "stAng")]
    StAng,
    #[serde(rename = "spanAng")]
    SpanAng,
    #[serde(rename = "ar")]
    Ar,
    #[serde(rename = "lnSpPar")]
    LnSpPar,
    #[serde(rename = "lnSpAfParP")]
    LnSpAfParP,
    #[serde(rename = "lnSpCh")]
    LnSpCh,
    #[serde(rename = "lnSpAfChP")]
    LnSpAfChP,
    #[serde(rename = "rtShortDist")]
    RtShortDist,
    #[serde(rename = "alignTx")]
    AlignTx,
    #[serde(rename = "pyraLvlNode")]
    PyraLvlNode,
    #[serde(rename = "pyraAcctBkgdNode")]
    PyraAcctBkgdNode,
    #[serde(rename = "pyraAcctTxNode")]
    PyraAcctTxNode,
    #[serde(rename = "srcNode")]
    SrcNode,
    #[serde(rename = "dstNode")]
    DstNode,
    #[serde(rename = "begPts")]
    BegPts,
    #[serde(rename = "endPts")]
    EndPts,
}

impl std::fmt::Display for STParameterId {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::HorzAlign => write!(f, "horzAlign"),
            Self::VertAlign => write!(f, "vertAlign"),
            Self::ChDir => write!(f, "chDir"),
            Self::ChAlign => write!(f, "chAlign"),
            Self::SecChAlign => write!(f, "secChAlign"),
            Self::LinDir => write!(f, "linDir"),
            Self::SecLinDir => write!(f, "secLinDir"),
            Self::StElem => write!(f, "stElem"),
            Self::BendPt => write!(f, "bendPt"),
            Self::ConnRout => write!(f, "connRout"),
            Self::BegSty => write!(f, "begSty"),
            Self::EndSty => write!(f, "endSty"),
            Self::Dim => write!(f, "dim"),
            Self::RotPath => write!(f, "rotPath"),
            Self::CtrShpMap => write!(f, "ctrShpMap"),
            Self::NodeHorzAlign => write!(f, "nodeHorzAlign"),
            Self::NodeVertAlign => write!(f, "nodeVertAlign"),
            Self::Fallback => write!(f, "fallback"),
            Self::TxDir => write!(f, "txDir"),
            Self::PyraAcctPos => write!(f, "pyraAcctPos"),
            Self::PyraAcctTxMar => write!(f, "pyraAcctTxMar"),
            Self::TxBlDir => write!(f, "txBlDir"),
            Self::TxAnchorHorz => write!(f, "txAnchorHorz"),
            Self::TxAnchorVert => write!(f, "txAnchorVert"),
            Self::TxAnchorHorzCh => write!(f, "txAnchorHorzCh"),
            Self::TxAnchorVertCh => write!(f, "txAnchorVertCh"),
            Self::ParTxLTRAlign => write!(f, "parTxLTRAlign"),
            Self::ParTxRTLAlign => write!(f, "parTxRTLAlign"),
            Self::ShpTxLTRAlignCh => write!(f, "shpTxLTRAlignCh"),
            Self::ShpTxRTLAlignCh => write!(f, "shpTxRTLAlignCh"),
            Self::AutoTxRot => write!(f, "autoTxRot"),
            Self::GrDir => write!(f, "grDir"),
            Self::FlowDir => write!(f, "flowDir"),
            Self::ContDir => write!(f, "contDir"),
            Self::Bkpt => write!(f, "bkpt"),
            Self::Off => write!(f, "off"),
            Self::HierAlign => write!(f, "hierAlign"),
            Self::BkPtFixedVal => write!(f, "bkPtFixedVal"),
            Self::StBulletLvl => write!(f, "stBulletLvl"),
            Self::StAng => write!(f, "stAng"),
            Self::SpanAng => write!(f, "spanAng"),
            Self::Ar => write!(f, "ar"),
            Self::LnSpPar => write!(f, "lnSpPar"),
            Self::LnSpAfParP => write!(f, "lnSpAfParP"),
            Self::LnSpCh => write!(f, "lnSpCh"),
            Self::LnSpAfChP => write!(f, "lnSpAfChP"),
            Self::RtShortDist => write!(f, "rtShortDist"),
            Self::AlignTx => write!(f, "alignTx"),
            Self::PyraLvlNode => write!(f, "pyraLvlNode"),
            Self::PyraAcctBkgdNode => write!(f, "pyraAcctBkgdNode"),
            Self::PyraAcctTxNode => write!(f, "pyraAcctTxNode"),
            Self::SrcNode => write!(f, "srcNode"),
            Self::DstNode => write!(f, "dstNode"),
            Self::BegPts => write!(f, "begPts"),
            Self::EndPts => write!(f, "endPts"),
        }
    }
}

impl std::str::FromStr for STParameterId {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "horzAlign" => Ok(Self::HorzAlign),
            "vertAlign" => Ok(Self::VertAlign),
            "chDir" => Ok(Self::ChDir),
            "chAlign" => Ok(Self::ChAlign),
            "secChAlign" => Ok(Self::SecChAlign),
            "linDir" => Ok(Self::LinDir),
            "secLinDir" => Ok(Self::SecLinDir),
            "stElem" => Ok(Self::StElem),
            "bendPt" => Ok(Self::BendPt),
            "connRout" => Ok(Self::ConnRout),
            "begSty" => Ok(Self::BegSty),
            "endSty" => Ok(Self::EndSty),
            "dim" => Ok(Self::Dim),
            "rotPath" => Ok(Self::RotPath),
            "ctrShpMap" => Ok(Self::CtrShpMap),
            "nodeHorzAlign" => Ok(Self::NodeHorzAlign),
            "nodeVertAlign" => Ok(Self::NodeVertAlign),
            "fallback" => Ok(Self::Fallback),
            "txDir" => Ok(Self::TxDir),
            "pyraAcctPos" => Ok(Self::PyraAcctPos),
            "pyraAcctTxMar" => Ok(Self::PyraAcctTxMar),
            "txBlDir" => Ok(Self::TxBlDir),
            "txAnchorHorz" => Ok(Self::TxAnchorHorz),
            "txAnchorVert" => Ok(Self::TxAnchorVert),
            "txAnchorHorzCh" => Ok(Self::TxAnchorHorzCh),
            "txAnchorVertCh" => Ok(Self::TxAnchorVertCh),
            "parTxLTRAlign" => Ok(Self::ParTxLTRAlign),
            "parTxRTLAlign" => Ok(Self::ParTxRTLAlign),
            "shpTxLTRAlignCh" => Ok(Self::ShpTxLTRAlignCh),
            "shpTxRTLAlignCh" => Ok(Self::ShpTxRTLAlignCh),
            "autoTxRot" => Ok(Self::AutoTxRot),
            "grDir" => Ok(Self::GrDir),
            "flowDir" => Ok(Self::FlowDir),
            "contDir" => Ok(Self::ContDir),
            "bkpt" => Ok(Self::Bkpt),
            "off" => Ok(Self::Off),
            "hierAlign" => Ok(Self::HierAlign),
            "bkPtFixedVal" => Ok(Self::BkPtFixedVal),
            "stBulletLvl" => Ok(Self::StBulletLvl),
            "stAng" => Ok(Self::StAng),
            "spanAng" => Ok(Self::SpanAng),
            "ar" => Ok(Self::Ar),
            "lnSpPar" => Ok(Self::LnSpPar),
            "lnSpAfParP" => Ok(Self::LnSpAfParP),
            "lnSpCh" => Ok(Self::LnSpCh),
            "lnSpAfChP" => Ok(Self::LnSpAfChP),
            "rtShortDist" => Ok(Self::RtShortDist),
            "alignTx" => Ok(Self::AlignTx),
            "pyraLvlNode" => Ok(Self::PyraLvlNode),
            "pyraAcctBkgdNode" => Ok(Self::PyraAcctBkgdNode),
            "pyraAcctTxNode" => Ok(Self::PyraAcctTxNode),
            "srcNode" => Ok(Self::SrcNode),
            "dstNode" => Ok(Self::DstNode),
            "begPts" => Ok(Self::BegPts),
            "endPts" => Ok(Self::EndPts),
            _ => Err(format!("unknown STParameterId value: {}", s)),
        }
    }
}

pub type STInts = String;

pub type STUnsignedInts = String;

pub type STBooleans = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFunctionType {
    #[serde(rename = "cnt")]
    Cnt,
    #[serde(rename = "pos")]
    Pos,
    #[serde(rename = "revPos")]
    RevPos,
    #[serde(rename = "posEven")]
    PosEven,
    #[serde(rename = "posOdd")]
    PosOdd,
    #[serde(rename = "var")]
    Var,
    #[serde(rename = "depth")]
    Depth,
    #[serde(rename = "maxDepth")]
    MaxDepth,
}

impl std::fmt::Display for STFunctionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cnt => write!(f, "cnt"),
            Self::Pos => write!(f, "pos"),
            Self::RevPos => write!(f, "revPos"),
            Self::PosEven => write!(f, "posEven"),
            Self::PosOdd => write!(f, "posOdd"),
            Self::Var => write!(f, "var"),
            Self::Depth => write!(f, "depth"),
            Self::MaxDepth => write!(f, "maxDepth"),
        }
    }
}

impl std::str::FromStr for STFunctionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cnt" => Ok(Self::Cnt),
            "pos" => Ok(Self::Pos),
            "revPos" => Ok(Self::RevPos),
            "posEven" => Ok(Self::PosEven),
            "posOdd" => Ok(Self::PosOdd),
            "var" => Ok(Self::Var),
            "depth" => Ok(Self::Depth),
            "maxDepth" => Ok(Self::MaxDepth),
            _ => Err(format!("unknown STFunctionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFunctionOperator {
    #[serde(rename = "equ")]
    Equ,
    #[serde(rename = "neq")]
    Neq,
    #[serde(rename = "gt")]
    Gt,
    #[serde(rename = "lt")]
    Lt,
    #[serde(rename = "gte")]
    Gte,
    #[serde(rename = "lte")]
    Lte,
}

impl std::fmt::Display for STFunctionOperator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Equ => write!(f, "equ"),
            Self::Neq => write!(f, "neq"),
            Self::Gt => write!(f, "gt"),
            Self::Lt => write!(f, "lt"),
            Self::Gte => write!(f, "gte"),
            Self::Lte => write!(f, "lte"),
        }
    }
}

impl std::str::FromStr for STFunctionOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "equ" => Ok(Self::Equ),
            "neq" => Ok(Self::Neq),
            "gt" => Ok(Self::Gt),
            "lt" => Ok(Self::Lt),
            "gte" => Ok(Self::Gte),
            "lte" => Ok(Self::Lte),
            _ => Err(format!("unknown STFunctionOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDiagramHorizontalAlignment {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STDiagramHorizontalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STDiagramHorizontalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STDiagramHorizontalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVerticalAlignment {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "mid")]
    Mid,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STVerticalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::Mid => write!(f, "mid"),
            Self::B => write!(f, "b"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STVerticalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "mid" => Ok(Self::Mid),
            "b" => Ok(Self::B),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STVerticalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STChildDirection {
    #[serde(rename = "horz")]
    Horz,
    #[serde(rename = "vert")]
    Vert,
}

impl std::fmt::Display for STChildDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Horz => write!(f, "horz"),
            Self::Vert => write!(f, "vert"),
        }
    }
}

impl std::str::FromStr for STChildDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "horz" => Ok(Self::Horz),
            "vert" => Ok(Self::Vert),
            _ => Err(format!("unknown STChildDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STChildAlignment {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
}

impl std::fmt::Display for STChildAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::B => write!(f, "b"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
        }
    }
}

impl std::str::FromStr for STChildAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "b" => Ok(Self::B),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            _ => Err(format!("unknown STChildAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSecondaryChildAlignment {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
}

impl std::fmt::Display for STSecondaryChildAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::T => write!(f, "t"),
            Self::B => write!(f, "b"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
        }
    }
}

impl std::str::FromStr for STSecondaryChildAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "t" => Ok(Self::T),
            "b" => Ok(Self::B),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            _ => Err(format!("unknown STSecondaryChildAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLinearDirection {
    #[serde(rename = "fromL")]
    FromL,
    #[serde(rename = "fromR")]
    FromR,
    #[serde(rename = "fromT")]
    FromT,
    #[serde(rename = "fromB")]
    FromB,
}

impl std::fmt::Display for STLinearDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FromL => write!(f, "fromL"),
            Self::FromR => write!(f, "fromR"),
            Self::FromT => write!(f, "fromT"),
            Self::FromB => write!(f, "fromB"),
        }
    }
}

impl std::str::FromStr for STLinearDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fromL" => Ok(Self::FromL),
            "fromR" => Ok(Self::FromR),
            "fromT" => Ok(Self::FromT),
            "fromB" => Ok(Self::FromB),
            _ => Err(format!("unknown STLinearDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSecondaryLinearDirection {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "fromL")]
    FromL,
    #[serde(rename = "fromR")]
    FromR,
    #[serde(rename = "fromT")]
    FromT,
    #[serde(rename = "fromB")]
    FromB,
}

impl std::fmt::Display for STSecondaryLinearDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::FromL => write!(f, "fromL"),
            Self::FromR => write!(f, "fromR"),
            Self::FromT => write!(f, "fromT"),
            Self::FromB => write!(f, "fromB"),
        }
    }
}

impl std::str::FromStr for STSecondaryLinearDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "fromL" => Ok(Self::FromL),
            "fromR" => Ok(Self::FromR),
            "fromT" => Ok(Self::FromT),
            "fromB" => Ok(Self::FromB),
            _ => Err(format!("unknown STSecondaryLinearDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STStartingElement {
    #[serde(rename = "node")]
    Node,
    #[serde(rename = "trans")]
    Trans,
}

impl std::fmt::Display for STStartingElement {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Node => write!(f, "node"),
            Self::Trans => write!(f, "trans"),
        }
    }
}

impl std::str::FromStr for STStartingElement {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "node" => Ok(Self::Node),
            "trans" => Ok(Self::Trans),
            _ => Err(format!("unknown STStartingElement value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRotationPath {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "alongPath")]
    AlongPath,
}

impl std::fmt::Display for STRotationPath {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::AlongPath => write!(f, "alongPath"),
        }
    }
}

impl std::str::FromStr for STRotationPath {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "alongPath" => Ok(Self::AlongPath),
            _ => Err(format!("unknown STRotationPath value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCenterShapeMapping {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "fNode")]
    FNode,
}

impl std::fmt::Display for STCenterShapeMapping {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::FNode => write!(f, "fNode"),
        }
    }
}

impl std::str::FromStr for STCenterShapeMapping {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "fNode" => Ok(Self::FNode),
            _ => Err(format!("unknown STCenterShapeMapping value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBendPoint {
    #[serde(rename = "beg")]
    Beg,
    #[serde(rename = "def")]
    Def,
    #[serde(rename = "end")]
    End,
}

impl std::fmt::Display for STBendPoint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Beg => write!(f, "beg"),
            Self::Def => write!(f, "def"),
            Self::End => write!(f, "end"),
        }
    }
}

impl std::str::FromStr for STBendPoint {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "beg" => Ok(Self::Beg),
            "def" => Ok(Self::Def),
            "end" => Ok(Self::End),
            _ => Err(format!("unknown STBendPoint value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConnectorRouting {
    #[serde(rename = "stra")]
    Stra,
    #[serde(rename = "bend")]
    Bend,
    #[serde(rename = "curve")]
    Curve,
    #[serde(rename = "longCurve")]
    LongCurve,
}

impl std::fmt::Display for STConnectorRouting {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stra => write!(f, "stra"),
            Self::Bend => write!(f, "bend"),
            Self::Curve => write!(f, "curve"),
            Self::LongCurve => write!(f, "longCurve"),
        }
    }
}

impl std::str::FromStr for STConnectorRouting {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stra" => Ok(Self::Stra),
            "bend" => Ok(Self::Bend),
            "curve" => Ok(Self::Curve),
            "longCurve" => Ok(Self::LongCurve),
            _ => Err(format!("unknown STConnectorRouting value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STArrowheadStyle {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "arr")]
    Arr,
    #[serde(rename = "noArr")]
    NoArr,
}

impl std::fmt::Display for STArrowheadStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::Arr => write!(f, "arr"),
            Self::NoArr => write!(f, "noArr"),
        }
    }
}

impl std::str::FromStr for STArrowheadStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "arr" => Ok(Self::Arr),
            "noArr" => Ok(Self::NoArr),
            _ => Err(format!("unknown STArrowheadStyle value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConnectorDimension {
    #[serde(rename = "1D")]
    _1D,
    #[serde(rename = "2D")]
    _2D,
    #[serde(rename = "cust")]
    Cust,
}

impl std::fmt::Display for STConnectorDimension {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::_1D => write!(f, "1D"),
            Self::_2D => write!(f, "2D"),
            Self::Cust => write!(f, "cust"),
        }
    }
}

impl std::str::FromStr for STConnectorDimension {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "1D" => Ok(Self::_1D),
            "2D" => Ok(Self::_2D),
            "cust" => Ok(Self::Cust),
            _ => Err(format!("unknown STConnectorDimension value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConnectorPoint {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "bCtr")]
    BCtr,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "midL")]
    MidL,
    #[serde(rename = "midR")]
    MidR,
    #[serde(rename = "tCtr")]
    TCtr,
    #[serde(rename = "bL")]
    BL,
    #[serde(rename = "bR")]
    BR,
    #[serde(rename = "tL")]
    TL,
    #[serde(rename = "tR")]
    TR,
    #[serde(rename = "radial")]
    Radial,
}

impl std::fmt::Display for STConnectorPoint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::BCtr => write!(f, "bCtr"),
            Self::Ctr => write!(f, "ctr"),
            Self::MidL => write!(f, "midL"),
            Self::MidR => write!(f, "midR"),
            Self::TCtr => write!(f, "tCtr"),
            Self::BL => write!(f, "bL"),
            Self::BR => write!(f, "bR"),
            Self::TL => write!(f, "tL"),
            Self::TR => write!(f, "tR"),
            Self::Radial => write!(f, "radial"),
        }
    }
}

impl std::str::FromStr for STConnectorPoint {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "bCtr" => Ok(Self::BCtr),
            "ctr" => Ok(Self::Ctr),
            "midL" => Ok(Self::MidL),
            "midR" => Ok(Self::MidR),
            "tCtr" => Ok(Self::TCtr),
            "bL" => Ok(Self::BL),
            "bR" => Ok(Self::BR),
            "tL" => Ok(Self::TL),
            "tR" => Ok(Self::TR),
            "radial" => Ok(Self::Radial),
            _ => Err(format!("unknown STConnectorPoint value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STNodeHorizontalAlignment {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
}

impl std::fmt::Display for STNodeHorizontalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
        }
    }
}

impl std::str::FromStr for STNodeHorizontalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            _ => Err(format!("unknown STNodeHorizontalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STNodeVerticalAlignment {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "mid")]
    Mid,
    #[serde(rename = "b")]
    B,
}

impl std::fmt::Display for STNodeVerticalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::Mid => write!(f, "mid"),
            Self::B => write!(f, "b"),
        }
    }
}

impl std::str::FromStr for STNodeVerticalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "mid" => Ok(Self::Mid),
            "b" => Ok(Self::B),
            _ => Err(format!("unknown STNodeVerticalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFallbackDimension {
    #[serde(rename = "1D")]
    _1D,
    #[serde(rename = "2D")]
    _2D,
}

impl std::fmt::Display for STFallbackDimension {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::_1D => write!(f, "1D"),
            Self::_2D => write!(f, "2D"),
        }
    }
}

impl std::str::FromStr for STFallbackDimension {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "1D" => Ok(Self::_1D),
            "2D" => Ok(Self::_2D),
            _ => Err(format!("unknown STFallbackDimension value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextDirection {
    #[serde(rename = "fromT")]
    FromT,
    #[serde(rename = "fromB")]
    FromB,
}

impl std::fmt::Display for STTextDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FromT => write!(f, "fromT"),
            Self::FromB => write!(f, "fromB"),
        }
    }
}

impl std::str::FromStr for STTextDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fromT" => Ok(Self::FromT),
            "fromB" => Ok(Self::FromB),
            _ => Err(format!("unknown STTextDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPyramidAccentPosition {
    #[serde(rename = "bef")]
    Bef,
    #[serde(rename = "aft")]
    Aft,
}

impl std::fmt::Display for STPyramidAccentPosition {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Bef => write!(f, "bef"),
            Self::Aft => write!(f, "aft"),
        }
    }
}

impl std::str::FromStr for STPyramidAccentPosition {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bef" => Ok(Self::Bef),
            "aft" => Ok(Self::Aft),
            _ => Err(format!("unknown STPyramidAccentPosition value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPyramidAccentTextMargin {
    #[serde(rename = "step")]
    Step,
    #[serde(rename = "stack")]
    Stack,
}

impl std::fmt::Display for STPyramidAccentTextMargin {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Step => write!(f, "step"),
            Self::Stack => write!(f, "stack"),
        }
    }
}

impl std::str::FromStr for STPyramidAccentTextMargin {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "step" => Ok(Self::Step),
            "stack" => Ok(Self::Stack),
            _ => Err(format!("unknown STPyramidAccentTextMargin value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextBlockDirection {
    #[serde(rename = "horz")]
    Horz,
    #[serde(rename = "vert")]
    Vert,
}

impl std::fmt::Display for STTextBlockDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Horz => write!(f, "horz"),
            Self::Vert => write!(f, "vert"),
        }
    }
}

impl std::str::FromStr for STTextBlockDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "horz" => Ok(Self::Horz),
            "vert" => Ok(Self::Vert),
            _ => Err(format!("unknown STTextBlockDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAnchorHorizontal {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "ctr")]
    Ctr,
}

impl std::fmt::Display for STTextAnchorHorizontal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Ctr => write!(f, "ctr"),
        }
    }
}

impl std::str::FromStr for STTextAnchorHorizontal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "ctr" => Ok(Self::Ctr),
            _ => Err(format!("unknown STTextAnchorHorizontal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAnchorVertical {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "mid")]
    Mid,
    #[serde(rename = "b")]
    B,
}

impl std::fmt::Display for STTextAnchorVertical {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::Mid => write!(f, "mid"),
            Self::B => write!(f, "b"),
        }
    }
}

impl std::str::FromStr for STTextAnchorVertical {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "mid" => Ok(Self::Mid),
            "b" => Ok(Self::B),
            _ => Err(format!("unknown STTextAnchorVertical value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDiagramTextAlignment {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
}

impl std::fmt::Display for STDiagramTextAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
        }
    }
}

impl std::str::FromStr for STDiagramTextAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            _ => Err(format!("unknown STDiagramTextAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAutoTextRotation {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "upr")]
    Upr,
    #[serde(rename = "grav")]
    Grav,
}

impl std::fmt::Display for STAutoTextRotation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Upr => write!(f, "upr"),
            Self::Grav => write!(f, "grav"),
        }
    }
}

impl std::str::FromStr for STAutoTextRotation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "upr" => Ok(Self::Upr),
            "grav" => Ok(Self::Grav),
            _ => Err(format!("unknown STAutoTextRotation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STGrowDirection {
    #[serde(rename = "tL")]
    TL,
    #[serde(rename = "tR")]
    TR,
    #[serde(rename = "bL")]
    BL,
    #[serde(rename = "bR")]
    BR,
}

impl std::fmt::Display for STGrowDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TL => write!(f, "tL"),
            Self::TR => write!(f, "tR"),
            Self::BL => write!(f, "bL"),
            Self::BR => write!(f, "bR"),
        }
    }
}

impl std::str::FromStr for STGrowDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tL" => Ok(Self::TL),
            "tR" => Ok(Self::TR),
            "bL" => Ok(Self::BL),
            "bR" => Ok(Self::BR),
            _ => Err(format!("unknown STGrowDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFlowDirection {
    #[serde(rename = "row")]
    Row,
    #[serde(rename = "col")]
    Col,
}

impl std::fmt::Display for STFlowDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Row => write!(f, "row"),
            Self::Col => write!(f, "col"),
        }
    }
}

impl std::str::FromStr for STFlowDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "row" => Ok(Self::Row),
            "col" => Ok(Self::Col),
            _ => Err(format!("unknown STFlowDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STContinueDirection {
    #[serde(rename = "revDir")]
    RevDir,
    #[serde(rename = "sameDir")]
    SameDir,
}

impl std::fmt::Display for STContinueDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RevDir => write!(f, "revDir"),
            Self::SameDir => write!(f, "sameDir"),
        }
    }
}

impl std::str::FromStr for STContinueDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "revDir" => Ok(Self::RevDir),
            "sameDir" => Ok(Self::SameDir),
            _ => Err(format!("unknown STContinueDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBreakpoint {
    #[serde(rename = "endCnv")]
    EndCnv,
    #[serde(rename = "bal")]
    Bal,
    #[serde(rename = "fixed")]
    Fixed,
}

impl std::fmt::Display for STBreakpoint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::EndCnv => write!(f, "endCnv"),
            Self::Bal => write!(f, "bal"),
            Self::Fixed => write!(f, "fixed"),
        }
    }
}

impl std::str::FromStr for STBreakpoint {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "endCnv" => Ok(Self::EndCnv),
            "bal" => Ok(Self::Bal),
            "fixed" => Ok(Self::Fixed),
            _ => Err(format!("unknown STBreakpoint value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOffset {
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "off")]
    Off,
}

impl std::fmt::Display for STOffset {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ctr => write!(f, "ctr"),
            Self::Off => write!(f, "off"),
        }
    }
}

impl std::str::FromStr for STOffset {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ctr" => Ok(Self::Ctr),
            "off" => Ok(Self::Off),
            _ => Err(format!("unknown STOffset value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHierarchyAlignment {
    #[serde(rename = "tL")]
    TL,
    #[serde(rename = "tR")]
    TR,
    #[serde(rename = "tCtrCh")]
    TCtrCh,
    #[serde(rename = "tCtrDes")]
    TCtrDes,
    #[serde(rename = "bL")]
    BL,
    #[serde(rename = "bR")]
    BR,
    #[serde(rename = "bCtrCh")]
    BCtrCh,
    #[serde(rename = "bCtrDes")]
    BCtrDes,
    #[serde(rename = "lT")]
    LT,
    #[serde(rename = "lB")]
    LB,
    #[serde(rename = "lCtrCh")]
    LCtrCh,
    #[serde(rename = "lCtrDes")]
    LCtrDes,
    #[serde(rename = "rT")]
    RT,
    #[serde(rename = "rB")]
    RB,
    #[serde(rename = "rCtrCh")]
    RCtrCh,
    #[serde(rename = "rCtrDes")]
    RCtrDes,
}

impl std::fmt::Display for STHierarchyAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TL => write!(f, "tL"),
            Self::TR => write!(f, "tR"),
            Self::TCtrCh => write!(f, "tCtrCh"),
            Self::TCtrDes => write!(f, "tCtrDes"),
            Self::BL => write!(f, "bL"),
            Self::BR => write!(f, "bR"),
            Self::BCtrCh => write!(f, "bCtrCh"),
            Self::BCtrDes => write!(f, "bCtrDes"),
            Self::LT => write!(f, "lT"),
            Self::LB => write!(f, "lB"),
            Self::LCtrCh => write!(f, "lCtrCh"),
            Self::LCtrDes => write!(f, "lCtrDes"),
            Self::RT => write!(f, "rT"),
            Self::RB => write!(f, "rB"),
            Self::RCtrCh => write!(f, "rCtrCh"),
            Self::RCtrDes => write!(f, "rCtrDes"),
        }
    }
}

impl std::str::FromStr for STHierarchyAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tL" => Ok(Self::TL),
            "tR" => Ok(Self::TR),
            "tCtrCh" => Ok(Self::TCtrCh),
            "tCtrDes" => Ok(Self::TCtrDes),
            "bL" => Ok(Self::BL),
            "bR" => Ok(Self::BR),
            "bCtrCh" => Ok(Self::BCtrCh),
            "bCtrDes" => Ok(Self::BCtrDes),
            "lT" => Ok(Self::LT),
            "lB" => Ok(Self::LB),
            "lCtrCh" => Ok(Self::LCtrCh),
            "lCtrDes" => Ok(Self::LCtrDes),
            "rT" => Ok(Self::RT),
            "rB" => Ok(Self::RB),
            "rCtrCh" => Ok(Self::RCtrCh),
            "rCtrDes" => Ok(Self::RCtrDes),
            _ => Err(format!("unknown STHierarchyAlignment value: {}", s)),
        }
    }
}

pub type STFunctionValue = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVariableType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "orgChart")]
    OrgChart,
    #[serde(rename = "chMax")]
    ChMax,
    #[serde(rename = "chPref")]
    ChPref,
    #[serde(rename = "bulEnabled")]
    BulEnabled,
    #[serde(rename = "dir")]
    Dir,
    #[serde(rename = "hierBranch")]
    HierBranch,
    #[serde(rename = "animOne")]
    AnimOne,
    #[serde(rename = "animLvl")]
    AnimLvl,
    #[serde(rename = "resizeHandles")]
    ResizeHandles,
}

impl std::fmt::Display for STVariableType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::OrgChart => write!(f, "orgChart"),
            Self::ChMax => write!(f, "chMax"),
            Self::ChPref => write!(f, "chPref"),
            Self::BulEnabled => write!(f, "bulEnabled"),
            Self::Dir => write!(f, "dir"),
            Self::HierBranch => write!(f, "hierBranch"),
            Self::AnimOne => write!(f, "animOne"),
            Self::AnimLvl => write!(f, "animLvl"),
            Self::ResizeHandles => write!(f, "resizeHandles"),
        }
    }
}

impl std::str::FromStr for STVariableType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "orgChart" => Ok(Self::OrgChart),
            "chMax" => Ok(Self::ChMax),
            "chPref" => Ok(Self::ChPref),
            "bulEnabled" => Ok(Self::BulEnabled),
            "dir" => Ok(Self::Dir),
            "hierBranch" => Ok(Self::HierBranch),
            "animOne" => Ok(Self::AnimOne),
            "animLvl" => Ok(Self::AnimLvl),
            "resizeHandles" => Ok(Self::ResizeHandles),
            _ => Err(format!("unknown STVariableType value: {}", s)),
        }
    }
}

pub type STFunctionArgument = STVariableType;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOutputShapeType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "conn")]
    Conn,
}

impl std::fmt::Display for STOutputShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Conn => write!(f, "conn"),
        }
    }
}

impl std::str::FromStr for STOutputShapeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "conn" => Ok(Self::Conn),
            _ => Err(format!("unknown STOutputShapeType value: {}", s)),
        }
    }
}

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
    Tint(Box<PositiveFixedPercentageElement>),
    #[serde(rename = "shade")]
    Shade(Box<PositiveFixedPercentageElement>),
    #[serde(rename = "comp")]
    Comp(Box<CTComplementTransform>),
    #[serde(rename = "inv")]
    Inv(Box<CTInverseTransform>),
    #[serde(rename = "gray")]
    Gray(Box<CTGrayscaleTransform>),
    #[serde(rename = "alpha")]
    Alpha(Box<PositiveFixedPercentageElement>),
    #[serde(rename = "alphaOff")]
    AlphaOff(Box<FixedPercentageElement>),
    #[serde(rename = "alphaMod")]
    AlphaMod(Box<PositivePercentageElement>),
    #[serde(rename = "hue")]
    Hue(Box<CTPositiveFixedAngle>),
    #[serde(rename = "hueOff")]
    HueOff(Box<CTAngle>),
    #[serde(rename = "hueMod")]
    HueMod(Box<PositivePercentageElement>),
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
    AlphaMod(AlphaModulateEffectElement),
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
    BuSzPct(Box<TextBulletSizePercentElement>),
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
    BuBlip(TextBlipBulletElement),
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
    #[cfg(feature = "dml-media")]
    #[serde(rename = "@r:link")]
    pub link: STRelationshipId,
    #[cfg(feature = "dml-media")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVideoFile {
    #[cfg(feature = "dml-media")]
    #[serde(rename = "@r:link")]
    pub link: STRelationshipId,
    #[cfg(feature = "dml-media")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTQuickTimeFile {
    #[serde(rename = "@r:link")]
    pub link: STRelationshipId,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type AVideoFile = Box<CTVideoFile>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColorScheme {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@name")]
    pub name: String,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "dk1")]
    pub dk1: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "lt1")]
    pub lt1: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "dk2")]
    pub dk2: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "lt2")]
    pub lt2: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent1")]
    pub accent1: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent2")]
    pub accent2: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent3")]
    pub accent3: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent4")]
    pub accent4: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent5")]
    pub accent5: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "accent6")]
    pub accent6: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "hlink")]
    pub hlink: Box<CTColor>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "folHlink")]
    pub fol_hlink: Box<CTColor>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCustomColor {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCustomColorList {
    #[serde(rename = "custClr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cust_clr: Vec<CTCustomColor>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub font: Vec<CTSupplementalFont>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEffectStyleItem {
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontScheme {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@name")]
    pub name: String,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "majorFont")]
    pub major_font: Box<CTFontCollection>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "minorFont")]
    pub minor_font: Box<CTFontCollection>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFillStyleList {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Vec<EGFillProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLineStyleList {
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub line: Vec<LineProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEffectStyleList {
    #[serde(rename = "effectStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effect_style: Vec<CTEffectStyleItem>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBackgroundFillStyleList {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Vec<EGFillProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStyleMatrix {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "fillStyleLst")]
    pub fill_style_lst: Box<CTFillStyleList>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "lnStyleLst")]
    pub ln_style_lst: Box<CTLineStyleList>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "effectStyleLst")]
    pub effect_style_lst: Box<CTEffectStyleList>,
    #[cfg(feature = "dml-themes")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBaseStyles {
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "clrScheme")]
    pub clr_scheme: Box<ColorScheme>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "fontScheme")]
    pub font_scheme: Box<FontScheme>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "fmtScheme")]
    pub fmt_scheme: Box<CTStyleMatrix>,
    #[cfg(feature = "dml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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

pub type OfficeArtExtensionAnyElement = String;

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
pub struct PositivePercentageElement {
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
pub struct FixedPercentageElement {
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
pub struct PositiveFixedPercentageElement {
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
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct SrgbColor {
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "@val")]
    pub value: HexColorRgb,
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct HslColor {
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "@hue")]
    pub hue: STPositiveFixedAngle,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "@sat")]
    pub sat: STPercentage,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "@lum")]
    pub lum: STPercentage,
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct SystemColor {
    #[serde(rename = "@val")]
    pub value: STSystemColorVal,
    #[serde(rename = "@lastClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_clr: Option<HexColorRgb>,
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct SchemeColor {
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "@val")]
    pub value: STSchemeColorVal,
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct PresetColor {
    #[serde(rename = "@val")]
    pub value: STPresetColorVal,
    #[serde(skip)]
    #[serde(default)]
    pub color_transform: Vec<EGColorTransform>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct EGOfficeArtExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extents: Vec<CTOfficeArtExtension>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOfficeArtExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extents: Vec<CTOfficeArtExtension>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
pub struct CTColor {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorMRU {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Vec<EGColorChoice>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AAGBlob {
    #[serde(rename = "@r:embed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed: Option<STRelationshipId>,
    #[serde(rename = "@r:link")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEmbeddedWAVAudioFile {
    #[serde(rename = "@r:embed")]
    pub embed: STRelationshipId,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTHyperlink {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@invalidUrl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invalid_url: Option<String>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@action")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub action: Option<String>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@tgtFrame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tgt_frame: Option<String>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@tooltip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@history")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub history: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@highlightClick")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub highlight_click: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@endSnd")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub end_snd: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "snd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snd: Option<Box<CTEmbeddedWAVAudioFile>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTConnectorLocking {
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTShapeLocking {
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPictureLocking {
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTContentPartLocking {
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    #[cfg(feature = "dml-text")]
    #[serde(rename = "hlinkClick")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_click: Option<Box<CTHyperlink>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "hlinkHover")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_hover: Option<Box<CTHyperlink>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTNonVisualConnectorProperties {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "cxnSpLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cxn_sp_locks: Option<Box<CTConnectorLocking>>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "stCxn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_cxn: Option<Box<CTConnection>>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "endCxn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_cxn: Option<Box<CTConnection>>,
    #[cfg(feature = "dml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTNonVisualPictureProperties {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "@preferRelativeResize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prefer_relative_resize: Option<bool>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "picLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pic_locks: Option<Box<CTPictureLocking>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTNonVisualGroupDrawingShapeProps {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "grpSpLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grp_sp_locks: Option<Box<CTGroupLocking>>,
    #[cfg(feature = "dml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTNonVisualGraphicFrameProperties {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "graphicFrameLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub graphic_frame_locks: Option<Box<CTGraphicalObjectFrameLocking>>,
    #[cfg(feature = "dml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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

pub type GraphicalObjectDataAnyElement = String;

pub type GraphicalObjectElement = Box<CTGraphicalObjectData>;

pub type AGraphic = GraphicalObjectElement;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
pub struct CTAnimationElementChoice {
    #[serde(rename = "dgm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dgm: Option<Box<CTAnimationDgmElement>>,
    #[serde(rename = "chart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart: Option<Box<CTAnimationChartElement>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
pub struct CTAnimationGraphicalObjectBuildProperties {
    #[serde(rename = "bldDgm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_dgm: Option<Box<CTAnimationDgmBuildProperties>>,
    #[serde(rename = "bldChart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_chart: Option<Box<CTAnimationChartBuildProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBackgroundFormatting {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTWholeE2oFormatting {
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGvmlUseShapeRectangle;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlTextShape {
    #[serde(rename = "txBody")]
    pub tx_body: Box<TextBody>,
    #[serde(rename = "useSpRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_sp_rect: Option<Box<CTGvmlUseShapeRectangle>>,
    #[serde(rename = "xfrm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Box<Transform2D>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGraphicalObjectFrame {
    #[serde(rename = "nvGraphicFramePr")]
    pub nv_graphic_frame_pr: Box<CTGvmlGraphicFrameNonVisual>,
    #[serde(rename = "graphic")]
    pub graphic: GraphicalObjectElement,
    #[serde(rename = "xfrm")]
    pub transform: Box<Transform2D>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGroupShape {
    #[serde(rename = "nvGrpSpPr")]
    pub nv_grp_sp_pr: Box<CTGvmlGroupShapeNonVisual>,
    #[serde(rename = "grpSpPr")]
    pub grp_sp_pr: Box<CTGroupShapeProperties>,
    #[serde(rename = "txSp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tx_sp: Vec<CTGvmlTextShape>,
    #[serde(rename = "sp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sp: Vec<CTGvmlShape>,
    #[serde(rename = "cxnSp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cxn_sp: Vec<CTGvmlConnector>,
    #[serde(rename = "pic")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pic: Vec<CTGvmlPicture>,
    #[serde(rename = "graphicFrame")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub graphic_frame: Vec<CTGvmlGraphicalObjectFrame>,
    #[serde(rename = "grpSp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grp_sp: Vec<CTGvmlGroupShape>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCamera {
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@prst")]
    pub preset: STPresetCameraType,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@fov")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fov: Option<STFOVAngle>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@zoom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom: Option<STPositivePercentage>,
    #[cfg(feature = "dml-3d")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLightRig {
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@rig")]
    pub rig: STLightRigType,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@dir")]
    pub dir: STLightRigDirection,
    #[cfg(feature = "dml-3d")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTScene3D {
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "camera")]
    pub camera: Box<CTCamera>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "lightRig")]
    pub light_rig: Box<CTLightRig>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "backdrop")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backdrop: Option<Box<CTBackdrop>>,
    #[cfg(feature = "dml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBevel {
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-3d")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTShape3D {
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@z")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z: Option<STCoordinate>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@extrusionH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extrusion_h: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@contourW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contour_w: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "@prstMaterial")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_material: Option<STPresetMaterialType>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "bevelT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bevel_t: Option<Box<CTBevel>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "bevelB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bevel_b: Option<Box<CTBevel>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "extrusionClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extrusion_clr: Option<Box<CTColor>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "contourClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contour_clr: Option<Box<CTColor>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
pub struct CTAlphaInverseEffect {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBlurEffect {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@rad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rad: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorReplaceEffect {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDuotoneEffect {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Vec<EGColorChoice>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGlowEffect {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@rad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rad: Option<STPositiveCoordinate>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGrayscaleEffect;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTInnerShadowEffect {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOuterShadowEffect {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@kx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<STFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@ky")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<STFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STRectAlignment>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPresetShadowEffect {
    #[serde(rename = "@prst")]
    pub preset: STPresetShadowVal,
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTReflectionEffect {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@stA")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_a: Option<STPositiveFixedPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@stPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_pos: Option<STPositiveFixedPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@endA")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_a: Option<STPositiveFixedPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@endPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_pos: Option<STPositiveFixedPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@fadeDir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fade_dir: Option<STPositiveFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@kx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<STFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@ky")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<STFixedAngle>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STRectAlignment>,
    #[cfg(feature = "dml-effects")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    #[cfg(feature = "dml-effects")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
pub struct SolidColorFill {
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGradientStop {
    #[serde(rename = "@pos")]
    pub pos: STPositiveFixedPercentage,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGradientStopList {
    #[serde(rename = "gs")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gs: Vec<CTGradientStop>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GradientFill {
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@flip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip: Option<STTileFlipMode>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "gsLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gs_lst: Option<Box<CTGradientStopList>>,
    #[serde(skip)]
    #[serde(default)]
    pub shade_properties: Option<Box<EGShadeProperties>>,
    #[cfg(feature = "dml-fills")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTStretchInfoProperties {
    #[serde(rename = "fillRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_rect: Option<Box<CTRelativeRect>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Blip {
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@r:embed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed: Option<STRelationshipId>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@r:link")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<STRelationshipId>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@cstate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cstate: Option<STBlipCompression>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaBiLevel")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_bi_level: Vec<CTAlphaBiLevelEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaCeiling")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_ceiling: Vec<CTAlphaCeilingEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaFloor")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_floor: Vec<CTAlphaFloorEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaInv")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_inv: Vec<CTAlphaInverseEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaMod")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_mod: Vec<AlphaModulateEffectElement>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaModFix")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_mod_fix: Vec<CTAlphaModulateFixedEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "alphaRepl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub alpha_repl: Vec<CTAlphaReplaceEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "biLevel")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bi_level: Vec<CTBiLevelEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "blur")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub blur: Vec<CTBlurEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "clrChange")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub clr_change: Vec<CTColorChangeEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "clrRepl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub clr_repl: Vec<CTColorReplaceEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "duotone")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub duotone: Vec<CTDuotoneEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "fillOverlay")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub fill_overlay: Vec<CTFillOverlayEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "grayscl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grayscl: Vec<CTGrayscaleEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "hsl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hsl: Vec<CTHSLEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "lum")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lum: Vec<CTLuminanceEffect>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "tint")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tint: Vec<CTTintEffect>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BlipFillProperties {
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@dpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dpi: Option<u32>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "blip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blip: Option<Box<Blip>>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "srcRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<Box<CTRelativeRect>>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_mode_properties: Option<Box<EGFillModeProperties>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct PatternFill {
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "@prst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset: Option<STPresetPatternVal>,
    #[cfg(feature = "dml-fills")]
    #[serde(rename = "fgClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_clr: Option<Box<CTColor>>,
    #[cfg(feature = "dml-fills")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGroupFillProperties;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFillProperties {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFillEffect {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFillOverlayEffect {
    #[serde(rename = "@blend")]
    pub blend: STBlendMode,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EffectContainer {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STEffectContainerType>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip)]
    #[serde(default)]
    pub effect: Vec<EGEffect>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

pub type AlphaModulateEffectElement = Box<EffectContainer>;

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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EffectList {
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "blur")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur: Option<Box<CTBlurEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "fillOverlay")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_overlay: Option<Box<CTFillOverlayEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "glow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glow: Option<Box<CTGlowEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "innerShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inner_shdw: Option<Box<CTInnerShadowEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "outerShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outer_shdw: Option<Box<CTOuterShadowEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "prstShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_shdw: Option<Box<CTPresetShadowEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "reflection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reflection: Option<Box<CTReflectionEffect>>,
    #[cfg(feature = "dml-effects")]
    #[serde(rename = "softEdge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub soft_edge: Option<Box<CTSoftEdgesEffect>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEffectProperties {
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGeomGuideList {
    #[serde(rename = "gd")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gd: Vec<CTGeomGuide>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAdjustHandleList {
    #[serde(rename = "ahXY")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ah_x_y: Vec<CTXYAdjustHandle>,
    #[serde(rename = "ahPolar")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ah_polar: Vec<CTPolarAdjustHandle>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTConnectionSiteList {
    #[serde(rename = "cxn")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cxn: Vec<CTConnectionSite>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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

pub type Path2DMoveToElement = Box<CTAdjPoint2D>;

pub type Path2DLineToElement = Box<CTAdjPoint2D>;

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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPath2DQuadBezierTo {
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<CTAdjPoint2D>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPath2DCubicBezierTo {
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<CTAdjPoint2D>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPath2DClose;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    #[serde(rename = "close")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub close: Vec<CTPath2DClose>,
    #[serde(rename = "moveTo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub move_to: Vec<Path2DMoveToElement>,
    #[serde(rename = "lnTo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ln_to: Vec<Path2DLineToElement>,
    #[serde(rename = "arcTo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub arc_to: Vec<CTPath2DArcTo>,
    #[serde(rename = "quadBezTo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub quad_bez_to: Vec<CTPath2DQuadBezierTo>,
    #[serde(rename = "cubicBezTo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cubic_bez_to: Vec<CTPath2DCubicBezierTo>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPath2DList {
    #[serde(rename = "path")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub path: Vec<CTPath2D>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLineEndProperties {
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STLineEndType>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STLineEndWidth>,
    #[cfg(feature = "dml-lines")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPresetLineDashProperties {
    #[cfg(feature = "dml-lines")]
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
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@d")]
    pub d: STPositivePercentage,
    #[cfg(feature = "dml-lines")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDashStopList {
    #[serde(rename = "ds")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ds: Vec<CTDashStop>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LineProperties {
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STLineWidth>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@cap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<STLineCap>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@cmpd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cmpd: Option<STCompoundLine>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STPenAlignment>,
    #[serde(skip)]
    #[serde(default)]
    pub line_fill_properties: Option<Box<EGLineFillProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub line_dash_properties: Option<Box<EGLineDashProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub line_join_properties: Option<Box<EGLineJoinProperties>>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "headEnd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end: Option<Box<CTLineEndProperties>>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "tailEnd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<Box<CTLineEndProperties>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTShapeProperties {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<STBlackWhiteMode>,
    #[serde(rename = "xfrm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Box<Transform2D>>,
    #[serde(skip)]
    #[serde(default)]
    pub geometry: Option<Box<EGGeometry>>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[cfg(feature = "dml-lines")]
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "sp3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp3d: Option<Box<CTShape3D>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGroupShapeProperties {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<STBlackWhiteMode>,
    #[serde(rename = "xfrm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Box<CTGroupTransform2D>>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStyleMatrixReference {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@idx")]
    pub idx: STStyleMatrixColumnIndex,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTFontReference {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@idx")]
    pub idx: STFontCollectionIndex,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct ShapeStyle {
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "lnRef")]
    pub ln_ref: Box<CTStyleMatrixReference>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "fillRef")]
    pub fill_ref: Box<CTStyleMatrixReference>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "effectRef")]
    pub effect_ref: Box<CTStyleMatrixReference>,
    #[cfg(feature = "dml-shapes")]
    #[serde(rename = "fontRef")]
    pub font_ref: Box<CTFontReference>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorMappingOverride {
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "masterClrMapping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub master_clr_mapping: Option<Box<CTEmptyElement>>,
    #[cfg(feature = "dml-colors")]
    #[serde(rename = "overrideClrMapping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub override_clr_mapping: Option<Box<CTColorMapping>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorSchemeList {
    #[serde(rename = "extraClrScheme")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extra_clr_scheme: Vec<CTColorSchemeAndMapping>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOfficeStyleSheet {
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "themeElements")]
    pub theme_elements: Box<CTBaseStyles>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "objectDefaults")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub object_defaults: Option<Box<CTObjectStyleDefaults>>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "extraClrSchemeLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extra_clr_scheme_lst: Option<Box<CTColorSchemeList>>,
    #[cfg(feature = "dml-themes")]
    #[serde(rename = "custClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_clr_lst: Option<Box<CTCustomColorList>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type ATheme = Box<CTOfficeStyleSheet>;

pub type AThemeOverride = Box<CTBaseStylesOverride>;

pub type AThemeManager = Box<CTEmptyElement>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableCellProperties {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@marL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<STCoordinate32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@marR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<STCoordinate32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@marT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_t: Option<STCoordinate32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@marB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_b: Option<STCoordinate32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@vert")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert: Option<STTextVerticalType>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@anchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<STTextAnchoringType>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@anchorCtr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anchor_ctr: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@horzOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_overflow: Option<STTextHorzOverflowType>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_l: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_r: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_t: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_b: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnTlToBr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_tl_to_br: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "lnBlToTr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_bl_to_tr: Option<Box<LineProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "cell3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell3_d: Option<Box<CTCell3D>>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[serde(rename = "headers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<Box<CTHeaders>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTHeaders {
    #[serde(rename = "header")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableGrid {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "gridCol")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_col: Vec<CTTableCol>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableCell {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@rowSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_span: Option<i32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@gridSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<i32>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@hMerge")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub h_merge: Option<bool>,
    #[cfg(feature = "dml-tables")]
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
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "txBody")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_body: Option<Box<TextBody>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_pr: Option<Box<CTTableCellProperties>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableRow {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@h")]
    pub height: STCoordinate,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tc: Vec<CTTableCell>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableProperties {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@rtl")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@firstRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_row: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@firstCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_col: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@lastRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub last_row: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@lastCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub last_col: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@bandRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub band_row: Option<bool>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "@bandCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub band_col: Option<bool>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    #[serde(rename = "tableStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_style: Option<Box<CTTableStyle>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tableStyleId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_style_id: Option<Guid>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTable {
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tblPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_pr: Option<Box<CTTableProperties>>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tblGrid")]
    pub tbl_grid: Box<CTTableGrid>,
    #[cfg(feature = "dml-tables")]
    #[serde(rename = "tr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tr: Vec<CTTableRow>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTThemeableLineStyle {
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(rename = "lnRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_ref: Option<Box<CTStyleMatrixReference>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableStyleTextStyle {
    #[serde(rename = "@b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b: Option<STOnOffStyleType>,
    #[serde(rename = "@i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<STOnOffStyleType>,
    #[serde(skip)]
    #[serde(default)]
    pub themeable_font_styles: Option<Box<EGThemeableFontStyles>>,
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Option<Box<EGColorChoice>>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableBackgroundStyle {
    #[serde(skip)]
    #[serde(default)]
    pub themeable_fill_style: Option<Box<EGThemeableFillStyle>>,
    #[serde(skip)]
    #[serde(default)]
    pub themeable_effect_style: Option<Box<EGThemeableEffectStyle>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableStyleCellStyle {
    #[serde(rename = "tcBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_bdr: Option<Box<CTTableCellBorderStyle>>,
    #[serde(skip)]
    #[serde(default)]
    pub themeable_fill_style: Option<Box<EGThemeableFillStyle>>,
    #[serde(rename = "cell3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell3_d: Option<Box<CTCell3D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableStyleList {
    #[serde(rename = "@def")]
    pub def: Guid,
    #[serde(rename = "tblStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tbl_style: Vec<CTTableStyle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

pub type ATblStyleLst = Box<CTTableStyleList>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextParagraph {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_run: Vec<EGTextRun>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "endParaRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_para_r_pr: Option<Box<TextCharacterProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBodyProperties {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<STAngle>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@spcFirstLastPara")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub spc_first_last_para: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@vertOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_overflow: Option<STTextVertOverflowType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@horzOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_overflow: Option<STTextHorzOverflowType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@vert")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert: Option<STTextVerticalType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@wrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<STTextWrappingType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@lIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub l_ins: Option<STCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@tIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t_ins: Option<STCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@rIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ins: Option<STCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@bIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_ins: Option<STCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@numCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_col: Option<STTextColumnCount>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@spcCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_col: Option<STPositiveCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@rtlCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl_col: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@fromWordArt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub from_word_art: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@anchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<STTextAnchoringType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@anchorCtr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anchor_ctr: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@forceAA")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub force_a_a: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@upright")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub upright: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@compatLnSpc")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compat_ln_spc: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "prstTxWarp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_tx_warp: Option<Box<CTPresetTextShape>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_autofit: Option<Box<EGTextAutofit>>,
    #[cfg(feature = "dml-3d")]
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[serde(skip)]
    #[serde(default)]
    pub text3_d: Option<Box<EGText3D>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextBody {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "bodyPr")]
    pub body_pr: Box<CTTextBodyProperties>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "lstStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lst_style: Option<Box<CTTextListStyle>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "p")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub p: Vec<TextParagraph>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBulletColorFollowText;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBulletSizeFollowText;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextBulletSizePercentElement {
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

pub type TextBlipBulletElement = Box<Blip>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextNoBullet;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextFont {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@typeface")]
    pub typeface: STTextTypeface,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@panose")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub panose: Option<Panose>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@pitchFamily")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch_family: Option<STPitchFamily>,
    #[cfg(feature = "dml-text")]
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
pub struct CTTextUnderlineFillGroupWrapper {
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextCharacterProperties {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@kumimoji")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub kumimoji: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Language>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@altLang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_lang: Option<Language>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sz: Option<STTextFontSize>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub u: Option<STTextUnderlineType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<STTextStrikeType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<STTextNonNegativePoint>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@cap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<STTextCapsType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@spc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc: Option<STTextPoint>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@normalizeH")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub normalize_h: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@baseline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline: Option<STPercentage>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@noProof")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_proof: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@dirty")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dirty: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@err")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub err: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@smtClean")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub smt_clean: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@smtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smt_id: Option<u32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@bmk")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bmk: Option<String>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub fill_properties: Option<Box<EGFillProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub effect_properties: Option<Box<EGEffectProperties>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTColor>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_underline_line: Option<Box<EGTextUnderlineLine>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_underline_fill: Option<Box<EGTextUnderlineFill>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "latin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latin: Option<Box<TextFont>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "ea")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ea: Option<Box<TextFont>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<TextFont>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "sym")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sym: Option<Box<TextFont>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "hlinkClick")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_click: Option<Box<CTHyperlink>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "hlinkMouseOver")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_mouse_over: Option<Box<CTHyperlink>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextTabStopList {
    #[serde(rename = "tab")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tab: Vec<CTTextTabStop>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextLineBreak {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<TextCharacterProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextSpacing {
    #[serde(rename = "spcPct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_pct: Option<Box<CTTextSpacingPercent>>,
    #[serde(rename = "spcPts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_pts: Option<Box<CTTextSpacingPoint>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextParagraphProperties {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@marL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<STTextMargin>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@marR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<STTextMargin>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@lvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl: Option<STTextIndentLevelType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@indent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<STTextIndent>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STTextAlignType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@defTabSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_tab_sz: Option<STCoordinate32>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@rtl")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@eaLnBrk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ea_ln_brk: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@fontAlgn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_algn: Option<STTextFontAlignType>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@latinLnBrk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub latin_ln_brk: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "@hangingPunct")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hanging_punct: Option<bool>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "lnSpc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_spc: Option<Box<CTTextSpacing>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "spcBef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_bef: Option<Box<CTTextSpacing>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "spcAft")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_aft: Option<Box<CTTextSpacing>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_bullet_color: Option<Box<EGTextBulletColor>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_bullet_size: Option<Box<EGTextBulletSize>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_bullet_typeface: Option<Box<EGTextBulletTypeface>>,
    #[serde(skip)]
    #[serde(default)]
    pub text_bullet: Option<Box<EGTextBullet>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "tabLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_lst: Option<Box<CTTextTabStopList>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "defRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_r_pr: Option<Box<TextCharacterProperties>>,
    #[cfg(feature = "dml-extensions")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextRun {
    #[cfg(feature = "dml-text")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<TextCharacterProperties>>,
    #[cfg(feature = "dml-text")]
    #[serde(rename = "t")]
    pub t: String,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDouble {
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
pub struct CTUnsignedInt {
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartRelId {
    #[cfg(feature = "dml-charts")]
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
pub struct ChartExtension {
    #[cfg(feature = "dml-charts")]
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
pub struct ChartExtensionList {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extents: Vec<ChartExtension>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NumericValue {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@idx")]
    pub idx: u32,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@formatCode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "v")]
    pub v: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct NumericData {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "formatCode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ptCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_count: Option<Box<CTUnsignedInt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<NumericValue>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NumericReference {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "f")]
    pub f: String,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numCache")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_cache: Option<Box<NumericData>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumericDataSource {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_ref: Option<Box<NumericReference>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numLit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_lit: Option<Box<NumericData>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StringValue {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@idx")]
    pub idx: u32,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "v")]
    pub v: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct StringData {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ptCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_count: Option<Box<CTUnsignedInt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<StringValue>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StringReference {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "f")]
    pub f: String,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "strCache")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub str_cache: Option<Box<StringData>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartText {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "strRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub str_ref: Option<Box<StringReference>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "rich")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rich: Option<Box<TextBody>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextLanguageId {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
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
pub struct MultiLevelStrLevel {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<StringValue>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MultiLevelStrData {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ptCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_count: Option<Box<CTUnsignedInt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lvl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lvl: Vec<MultiLevelStrLevel>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MultiLevelStrRef {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "f")]
    pub f: String,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "multiLvlStrCache")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_lvl_str_cache: Option<Box<MultiLevelStrData>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AxisDataSource {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "multiLvlStrRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_lvl_str_ref: Option<Box<MultiLevelStrRef>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_ref: Option<Box<NumericReference>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numLit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_lit: Option<Box<NumericData>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "strRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub str_ref: Option<Box<StringReference>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "strLit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub str_lit: Option<Box<StringData>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SeriesText {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "strRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub str_ref: Option<Box<StringReference>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "v")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v: Option<XmlString>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutTarget {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<LayoutTargetType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutMode {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<LayoutModeType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ManualLayout {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layoutTarget")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<Box<LayoutTarget>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "xMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_mode: Option<Box<LayoutMode>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "yMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_mode: Option<Box<LayoutMode>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "wMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_mode: Option<Box<LayoutMode>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "hMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_mode: Option<Box<LayoutMode>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "y")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartLayout {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "manualLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_layout: Option<Box<ManualLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartTitle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<ChartText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "overlay")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overlay: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RotX {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<RotXValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HPercent {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<HPercentValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RotY {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<RotYValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DepthPercent {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<DepthPercentValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Perspective {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<PerspectiveValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct View3D {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "rotX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot_x: Option<Box<RotX>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "hPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_percent: Option<Box<HPercent>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "rotY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot_y: Option<Box<RotY>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "depthPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub depth_percent: Option<Box<DepthPercent>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "rAngAx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ang_ax: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "perspective")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub perspective: Option<Box<Perspective>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartSurface {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "thickness")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub thickness: Option<Box<ChartThickness>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_options: Option<Box<PictureOptions>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartThickness {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: ChartThicknessValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DataTable {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showHorzBorder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_horz_border: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showVertBorder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_vert_border: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showOutline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_outline: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showKeys")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_keys: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GapAmount {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<GapAmountValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OverlapAmount {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<OverlapValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BubbleScale {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<BubbleScaleValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SizeRepresents {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<SizeRepresentsType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FirstSliceAngle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<FirstSliceAngValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HoleSize {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<HoleSizeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SplitType {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<SplitTypeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomSplit {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "secondPiePt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub second_pie_pt: Vec<CTUnsignedInt>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SecondPieSize {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<SecondPieSizeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartNumFmt {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@formatCode")]
    pub format_code: XmlString,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@sourceLinked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub source_linked: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LabelAlignment {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: LabelAlignType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataLabelPosition {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: DataLabelPositionType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGDLblShared {
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[serde(rename = "dLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl_pos: Option<Box<DataLabelPosition>>,
    #[serde(rename = "showLegendKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<Box<CTBoolean>>,
    #[serde(rename = "showVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<Box<CTBoolean>>,
    #[serde(rename = "showCatName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showSerName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<Box<CTBoolean>>,
    #[serde(rename = "showBubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<Box<CTBoolean>>,
    #[serde(rename = "separator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DchrtGroupDLbl {
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<ChartText>>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[serde(rename = "dLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl_pos: Option<Box<DataLabelPosition>>,
    #[serde(rename = "showLegendKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<Box<CTBoolean>>,
    #[serde(rename = "showVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<Box<CTBoolean>>,
    #[serde(rename = "showCatName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showSerName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<Box<CTBoolean>>,
    #[serde(rename = "showBubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<Box<CTBoolean>>,
    #[serde(rename = "separator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataLabel {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<ChartText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl_pos: Option<Box<DataLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showLegendKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showCatName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showSerName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showBubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "separator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DchrtGroupDLbls {
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[serde(rename = "dLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl_pos: Option<Box<DataLabelPosition>>,
    #[serde(rename = "showLegendKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<Box<CTBoolean>>,
    #[serde(rename = "showVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<Box<CTBoolean>>,
    #[serde(rename = "showCatName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showSerName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<Box<CTBoolean>>,
    #[serde(rename = "showPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<Box<CTBoolean>>,
    #[serde(rename = "showBubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<Box<CTBoolean>>,
    #[serde(rename = "separator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[serde(rename = "showLeaderLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_leader_lines: Option<Box<CTBoolean>>,
    #[serde(rename = "leaderLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_lines: Option<Box<ChartLines>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DataLabels {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_lbl: Vec<DataLabel>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl_pos: Option<Box<DataLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showLegendKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showCatName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showSerName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showPercent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showBubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "separator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showLeaderLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_leader_lines: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "leaderLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartMarkerStyle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: MarkerStyleType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartMarkerSize {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<MarkerSizeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartMarker {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "symbol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub symbol: Option<Box<ChartMarkerStyle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "size")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<ChartMarkerSize>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataPoint {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "invertIfNegative")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invert_if_negative: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<ChartMarker>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubble3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble3_d: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "explosion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<Box<CTUnsignedInt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_options: Option<Box<PictureOptions>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TrendlineType {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TrendlineTypeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TrendlineOrder {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TrendlineOrderValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TrendlinePeriod {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TrendlinePeriodValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TrendlineLabel {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<ChartText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Trendline {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendlineType")]
    pub trendline_type: Box<TrendlineType>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub order: Option<Box<TrendlineOrder>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "period")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub period: Option<Box<TrendlinePeriod>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "forward")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forward: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "backward")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backward: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "intercept")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub intercept: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dispRSqr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_r_sqr: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dispEq")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_eq: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendlineLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub trendline_lbl: Option<Box<TrendlineLabel>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ErrorDirection {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: ErrorDirectionType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ErrorBarType {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<ErrorBarTypeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ErrorValueType {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<ErrorValueTypeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ErrorBars {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errDir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_dir: Option<Box<ErrorDirection>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBarType")]
    pub err_bar_type: Box<ErrorBarType>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errValType")]
    pub err_val_type: Box<ErrorValueType>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "noEndCap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_end_cap: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "plus")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plus: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minus")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minus: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct UpDownBar {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct UpDownBars {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "upBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub up_bars: Option<Box<UpDownBar>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "downBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub down_bars: Option<Box<UpDownBar>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGSerShared {
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LineSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<ChartMarker>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendline")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub trendline: Vec<Trendline>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_bars: Option<Box<ErrorBars>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "smooth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScatterSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<ChartMarker>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendline")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub trendline: Vec<Trendline>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBars")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub err_bars: Vec<ErrorBars>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "xVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_val: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "yVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_val: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "smooth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RadarSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<ChartMarker>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BarSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "invertIfNegative")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invert_if_negative: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_options: Option<Box<PictureOptions>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendline")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub trendline: Vec<Trendline>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_bars: Option<Box<ErrorBars>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AreaSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_options: Option<Box<PictureOptions>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendline")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub trendline: Vec<Trendline>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBars")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub err_bars: Vec<ErrorBars>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PieSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "explosion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<Box<CTUnsignedInt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BubbleSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "invertIfNegative")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invert_if_negative: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dPt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d_pt: Vec<DataPoint>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "trendline")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub trendline: Vec<Trendline>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "errBars")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub err_bars: Vec<ErrorBars>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "xVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_val: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "yVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_val: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubbleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubble3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble3_d: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SurfaceSeries {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "order")]
    pub order: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<SeriesText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat: Option<Box<AxisDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<NumericDataSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartGrouping {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<GroupingType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartLines {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGLineChartShared {
    #[serde(rename = "grouping")]
    pub grouping: Box<ChartGrouping>,
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<LineSeries>,
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LineChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    pub grouping: Box<ChartGrouping>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<LineSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "hiLowLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hi_low_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "upDownBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub up_down_bars: Option<Box<UpDownBars>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "smooth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Line3DChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    pub grouping: Box<ChartGrouping>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<LineSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapDepth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_depth: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct StockChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<LineSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "hiLowLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hi_low_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "upDownBars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub up_down_bars: Option<Box<UpDownBars>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ScatterStyle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<ScatterStyleType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScatterChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scatterStyle")]
    pub scatter_style: Box<ScatterStyle>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<ScatterSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RadarStyle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<RadarStyleType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RadarChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "radarStyle")]
    pub radar_style: Box<RadarStyle>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<RadarSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BarGrouping {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<BarGroupingType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BarDirection {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<BarDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramShape {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<BarShapeType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGBarChartShared {
    #[serde(rename = "barDir")]
    pub bar_dir: Box<BarDirection>,
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<BarGrouping>>,
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<BarSeries>,
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BarChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "barDir")]
    pub bar_dir: Box<BarDirection>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<BarGrouping>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<BarSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "overlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overlap: Option<Box<OverlapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "serLines")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser_lines: Vec<ChartLines>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Bar3DChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "barDir")]
    pub bar_dir: Box<BarDirection>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<BarGrouping>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<BarSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapDepth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_depth: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGAreaChartShared {
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<ChartGrouping>>,
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<AreaSeries>,
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AreaChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<ChartGrouping>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<AreaSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Area3DChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "grouping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<Box<ChartGrouping>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<AreaSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dropLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapDepth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_depth: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGPieChartShared {
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<PieSeries>,
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PieChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<PieSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "firstSliceAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_slice_ang: Option<Box<FirstSliceAngle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Pie3DChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<PieSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DoughnutChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<PieSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "firstSliceAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_slice_ang: Option<Box<FirstSliceAngle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "holeSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<Box<HoleSize>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OfPieType {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<OfPieTypeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OfPieChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ofPieType")]
    pub of_pie_type: Box<OfPieType>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<PieSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "gapWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<Box<GapAmount>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "splitType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_type: Option<Box<SplitType>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "splitPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_pos: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "custSplit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_split: Option<Box<CustomSplit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "secondPieSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub second_pie_size: Option<Box<SecondPieSize>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "serLines")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser_lines: Vec<ChartLines>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BubbleChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "varyColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<BubbleSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbls: Option<Box<DataLabels>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubble3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble3_d: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubbleScale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_scale: Option<Box<BubbleScale>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showNegBubbles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_neg_bubbles: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "sizeRepresents")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_represents: Option<Box<SizeRepresents>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BandFormat {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BandFormats {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bandFmt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub band_fmt: Vec<BandFormat>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGSurfaceChartShared {
    #[serde(rename = "wireframe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<Box<CTBoolean>>,
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<SurfaceSeries>,
    #[serde(rename = "bandFmts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band_fmts: Option<Box<BandFormats>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SurfaceChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "wireframe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<SurfaceSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bandFmts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band_fmts: Option<Box<BandFormats>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Surface3DChart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "wireframe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ser")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser: Vec<SurfaceSeries>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bandFmts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band_fmts: Option<Box<BandFormats>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ax_id: Vec<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AxisPosition {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: AxisPositionType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AxisCrosses {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: AxisCrossesType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CrossBetween {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: CrossBetweenType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TickMark {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TickMarkType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TickLabelPosition {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TickLabelPositionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AxisSkip {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: AxisSkipValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TimeUnit {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<TimeUnitType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AxisUnit {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: AxisUnitValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BuiltInUnit {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<BuiltInUnitType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartPictureFormat {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: PictureFormatType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PictureStackUnit {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: PictureStackUnitValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PictureOptions {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "applyToFront")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_front: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "applyToSides")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_sides: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "applyToEnd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_end: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_format: Option<Box<ChartPictureFormat>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pictureStackUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_stack_unit: Option<Box<PictureStackUnit>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DisplayUnitsLabel {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<Box<ChartText>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DisplayUnits {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "custUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_unit: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "builtInUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub built_in_unit: Option<Box<BuiltInUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dispUnitsLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_units_lbl: Option<Box<DisplayUnitsLabel>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AxisOrientation {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<AxisOrientationType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LogBase {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: LogBaseValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AxisScaling {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "logBase")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub log_base: Option<Box<LogBase>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "orientation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<Box<AxisOrientation>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "max")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "min")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LabelOffset {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<LabelOffsetValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGAxShared {
    #[serde(rename = "axId")]
    pub ax_id: Box<CTUnsignedInt>,
    #[serde(rename = "scaling")]
    pub scaling: Box<AxisScaling>,
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[serde(rename = "axPos")]
    pub ax_pos: Box<AxisPosition>,
    #[serde(rename = "majorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines: Option<Box<ChartLines>>,
    #[serde(rename = "minorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines: Option<Box<ChartLines>>,
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[serde(rename = "majorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<Box<TickMark>>,
    #[serde(rename = "minorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<Box<TickMark>>,
    #[serde(rename = "tickLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<Box<TickLabelPosition>>,
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[serde(rename = "crossAx")]
    pub cross_ax: Box<CTUnsignedInt>,
    #[serde(rename = "crosses")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<Box<AxisCrosses>>,
    #[serde(rename = "crossesAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<Box<CTDouble>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CategoryAxis {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    pub ax_id: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scaling")]
    pub scaling: Box<AxisScaling>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axPos")]
    pub ax_pos: Box<AxisPosition>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<Box<TickLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossAx")]
    pub cross_ax: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crosses")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<Box<AxisCrosses>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossesAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "auto")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lblAlgn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lbl_algn: Option<Box<LabelAlignment>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lblOffset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lbl_offset: Option<Box<LabelOffset>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblSkip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_skip: Option<Box<AxisSkip>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickMarkSkip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_mark_skip: Option<Box<AxisSkip>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "noMultiLvlLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_multi_lvl_lbl: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DateAxis {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    pub ax_id: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scaling")]
    pub scaling: Box<AxisScaling>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axPos")]
    pub ax_pos: Box<AxisPosition>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<Box<TickLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossAx")]
    pub cross_ax: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crosses")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<Box<AxisCrosses>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossesAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "auto")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lblOffset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lbl_offset: Option<Box<LabelOffset>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "baseTimeUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_time_unit: Option<Box<TimeUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<Box<AxisUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorTimeUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_time_unit: Option<Box<TimeUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<Box<AxisUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorTimeUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_time_unit: Option<Box<TimeUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SeriesAxis {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    pub ax_id: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scaling")]
    pub scaling: Box<AxisScaling>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axPos")]
    pub ax_pos: Box<AxisPosition>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<Box<TickLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossAx")]
    pub cross_ax: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crosses")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<Box<AxisCrosses>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossesAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblSkip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_skip: Option<Box<AxisSkip>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickMarkSkip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_mark_skip: Option<Box<AxisSkip>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValueAxis {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axId")]
    pub ax_id: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scaling")]
    pub scaling: Box<AxisScaling>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "axPos")]
    pub ax_pos: Box<AxisPosition>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorGridlines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines: Option<Box<ChartLines>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<ChartNumFmt>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorTickMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<Box<TickMark>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "tickLblPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<Box<TickLabelPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossAx")]
    pub cross_ax: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crosses")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<Box<AxisCrosses>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossesAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<Box<CTDouble>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "crossBetween")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_between: Option<Box<CrossBetween>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "majorUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<Box<AxisUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "minorUnit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<Box<AxisUnit>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dispUnits")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_units: Option<Box<DisplayUnits>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PlotArea {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "areaChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub area_chart: Vec<AreaChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "area3DChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub area3_d_chart: Vec<Area3DChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lineChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub line_chart: Vec<LineChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "line3DChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub line3_d_chart: Vec<Line3DChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "stockChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub stock_chart: Vec<StockChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "radarChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub radar_chart: Vec<RadarChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "scatterChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub scatter_chart: Vec<ScatterChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pieChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pie_chart: Vec<PieChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pie3DChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pie3_d_chart: Vec<Pie3DChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "doughnutChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub doughnut_chart: Vec<DoughnutChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "barChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bar_chart: Vec<BarChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bar3DChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bar3_d_chart: Vec<Bar3DChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "ofPieChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub of_pie_chart: Vec<OfPieChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "surfaceChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub surface_chart: Vec<SurfaceChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "surface3DChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub surface3_d_chart: Vec<Surface3DChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "bubbleChart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bubble_chart: Vec<BubbleChart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "valAx")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub val_ax: Vec<ValueAxis>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "catAx")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cat_ax: Vec<CategoryAxis>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dateAx")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub date_ax: Vec<DateAxis>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "serAx")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ser_ax: Vec<SeriesAxis>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dTable")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_table: Option<Box<DataTable>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotFormat {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "marker")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Box<ChartMarker>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub d_lbl: Option<Box<DataLabel>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotFormats {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pivotFmt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_fmt: Vec<PivotFormat>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LegendPosition {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<LegendPositionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGLegendEntryData {
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LegendEntry {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "idx")]
    pub idx: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "delete")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delete: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Legend {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "legendPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_pos: Option<Box<LegendPosition>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "legendEntry")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub legend_entry: Vec<LegendEntry>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<Box<ChartLayout>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "overlay")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overlay: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DisplayBlanksAs {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<DisplayBlanksAsType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Chart {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<ChartTitle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "autoTitleDeleted")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_title_deleted: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pivotFmts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_fmts: Option<Box<PivotFormats>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "view3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view3_d: Option<Box<View3D>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "floor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub floor: Option<Box<ChartSurface>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "sideWall")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub side_wall: Option<Box<ChartSurface>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "backWall")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_wall: Option<Box<ChartSurface>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "plotArea")]
    pub plot_area: Box<PlotArea>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "legend")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend: Option<Box<Legend>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "plotVisOnly")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_vis_only: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "dispBlanksAs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_blanks_as: Option<Box<DisplayBlanksAs>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "showDLblsOverMax")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_d_lbls_over_max: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartStyle {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@val")]
    pub value: ChartStyleValue,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotSource {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "name")]
    pub name: XmlString,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "fmtId")]
    pub fmt_id: Box<CTUnsignedInt>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ext_lst: Vec<ChartExtensionList>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartProtection {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "chartObject")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_object: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "data")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "formatting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formatting: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "selection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "userInterface")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user_interface: Option<Box<CTBoolean>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartHeaderFooter {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@alignWithMargins")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub align_with_margins: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@differentOddEven")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub different_odd_even: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@differentFirst")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub different_first: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "oddHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "oddFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "evenHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_header: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "evenFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "firstHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header: Option<XmlString>,
    #[cfg(feature = "dml-charts")]
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartPageMargins {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@l")]
    pub l: f64,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@r")]
    pub relationship_id: f64,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@t")]
    pub t: f64,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@b")]
    pub b: f64,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@header")]
    pub header: f64,
    #[cfg(feature = "dml-charts")]
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalData {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "autoUpdate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_update: Option<Box<CTBoolean>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct ChartPageSetup {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@paperSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@paperHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_height: Option<STPositiveUniversalMeasure>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@paperWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_width: Option<STPositiveUniversalMeasure>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@firstPageNumber")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_page_number: Option<u32>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@orientation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<ChartPageOrientation>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@blackAndWhite")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub black_and_white: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@draft")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub draft: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_first_page_number: Option<bool>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@horizontalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_dpi: Option<i32>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@verticalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_dpi: Option<i32>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "@copies")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copies: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PrintSettings {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<ChartHeaderFooter>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<ChartPageMargins>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<ChartPageSetup>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing_h_f: Option<Box<ChartRelId>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartSpace {
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "date1904")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date1904: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<TextLanguageId>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "roundedCorners")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rounded_corners: Option<Box<CTBoolean>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ChartStyle>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "clrMapOvr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_map_ovr: Option<Box<CTColorMapping>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "pivotSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_source: Option<Box<PivotSource>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "protection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Box<ChartProtection>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "chart")]
    pub chart: Box<Chart>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<TextBody>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "externalData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_data: Option<Box<ExternalData>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "printSettings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_settings: Option<Box<PrintSettings>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "userShapes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user_shapes: Option<Box<ChartRelId>>,
    #[cfg(feature = "dml-charts")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<ChartExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DchrtChartSpace = Box<ChartSpace>;

pub type DchrtUserShapes = String;

pub type DchrtChart = Box<ChartRelId>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColorTransformName {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct ColorTransformDescription {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct DiagramColorCategory {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@pri")]
    pub pri: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramColorCategories {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cat: Vec<DiagramColorCategory>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramColors {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@meth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub meth: Option<STClrAppMethod>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@hueDir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hue_dir: Option<STHueDir>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(skip)]
    #[serde(default)]
    pub color_choice: Vec<EGColorChoice>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct DiagramColorStyleLabel {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    pub name: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "fillClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "linClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lin_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "effectClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "txLinClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_lin_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "txFillClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_fill_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "txEffectClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_effect_clr_lst: Option<Box<DiagramColors>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramColorTransform {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<ColorTransformName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<ColorTransformDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramColorCategories>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "styleLbl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_lbl: Vec<DiagramColorStyleLabel>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmColorsDef = Box<DiagramColorTransform>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramColorTransformHeader {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    pub unique_id: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@resId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub res_id: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<ColorTransformName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<ColorTransformDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramColorCategories>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmColorsDefHdr = Box<DiagramColorTransformHeader>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramColorTransformHeaderList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "colorsDefHdr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub colors_def_hdr: Vec<DiagramColorTransformHeader>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmColorsDefHdrLst = Box<DiagramColorTransformHeaderList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramPoint {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@modelId")]
    pub model_id: STModelId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STPtType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@cxnId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cxn_id: Option<STModelId>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "prSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pr_set: Option<Box<DiagramElementProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "spPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_pr: Option<Box<CTShapeProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<Box<TextBody>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramPointList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<DiagramPoint>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramConnection {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@modelId")]
    pub model_id: STModelId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STCxnType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@srcId")]
    pub src_id: STModelId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@destId")]
    pub dest_id: STModelId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@srcOrd")]
    pub src_ord: u32,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@destOrd")]
    pub dest_ord: u32,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@parTransId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub par_trans_id: Option<STModelId>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@sibTransId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sib_trans_id: Option<STModelId>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramConnectionList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "cxn")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cxn: Vec<DiagramConnection>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataModel {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "ptLst")]
    pub pt_lst: Box<DiagramPointList>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "cxnLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cxn_lst: Option<Box<DiagramConnectionList>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "bg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg: Option<Box<CTBackgroundFormatting>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "whole")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub whole: Option<Box<CTWholeE2oFormatting>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmDataModel = Box<DataModel>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DdgrmAGIteratorAttributes {
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxisTypes>,
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementTypes>,
    #[serde(rename = "@hideLastTrans")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_last_trans: Option<STBooleans>,
    #[serde(rename = "@st")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st: Option<STInts>,
    #[serde(rename = "@cnt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnt: Option<STUnsignedInts>,
    #[serde(rename = "@step")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step: Option<STInts>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DdgrmAGConstraintAttributes {
    #[serde(rename = "@type")]
    pub r#type: STConstraintType,
    #[serde(rename = "@for")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#for: Option<STConstraintRelationship>,
    #[serde(rename = "@forName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub for_name: Option<String>,
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DdgrmAGConstraintRefAttributes {
    #[serde(rename = "@refType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_type: Option<STConstraintType>,
    #[serde(rename = "@refFor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_for: Option<STConstraintRelationship>,
    #[serde(rename = "@refForName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_for_name: Option<String>,
    #[serde(rename = "@refPtType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_pt_type: Option<STElementType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LayoutConstraint {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: STConstraintType,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@for")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#for: Option<STConstraintRelationship>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@forName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub for_name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@refType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_type: Option<STConstraintType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@refFor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_for: Option<STConstraintRelationship>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@refForName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_for_name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@refPtType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_pt_type: Option<STElementType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@op")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub op: Option<STBoolOperator>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@fact")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fact: Option<f64>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutConstraints {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "constr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub constr: Vec<LayoutConstraint>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NumericRule {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: STConstraintType,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@for")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#for: Option<STConstraintRelationship>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@forName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub for_name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@fact")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fact: Option<f64>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@max")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutRules {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "rule")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rule: Vec<NumericRule>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PresentationOf {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxisTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@hideLastTrans")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_last_trans: Option<STBooleans>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@st")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@cnt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnt: Option<STUnsignedInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@step")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LayoutAdjustment {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@idx")]
    pub idx: STIndex1,
    #[cfg(feature = "dml-diagrams")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutAdjustmentList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "adj")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub adj: Vec<LayoutAdjustment>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AlgorithmParameter {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: STParameterId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    pub value: STParameterVal,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LayoutAlgorithm {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: STAlgorithmType,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@rev")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rev: Option<u32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "param")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub param: Vec<AlgorithmParameter>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutNode {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@styleLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_lbl: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@chOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ch_order: Option<STChildOrderType>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@moveWith")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_with: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "alg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg: Option<Box<LayoutAlgorithm>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "presOf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_of: Option<Box<PresentationOf>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "constrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub constr_lst: Option<Box<LayoutConstraints>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "ruleLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule_lst: Option<Box<LayoutRules>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "varLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub var_lst: Option<Box<LayoutVariableProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "forEach")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub for_each: Vec<LayoutForEach>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutNode")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout_node: Vec<LayoutNode>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "choose")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub choose: Vec<LayoutChoose>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutForEach {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#ref: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxisTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@hideLastTrans")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_last_trans: Option<STBooleans>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@st")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@cnt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnt: Option<STUnsignedInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@step")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "alg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg: Option<Box<LayoutAlgorithm>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "presOf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_of: Option<Box<PresentationOf>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "constrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub constr_lst: Option<Box<LayoutConstraints>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "ruleLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule_lst: Option<Box<LayoutRules>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "forEach")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub for_each: Vec<LayoutForEach>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutNode")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout_node: Vec<LayoutNode>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "choose")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub choose: Vec<LayoutChoose>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LayoutWhen {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxisTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@ptType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pt_type: Option<STElementTypes>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@hideLastTrans")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_last_trans: Option<STBooleans>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@st")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@cnt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnt: Option<STUnsignedInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@step")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step: Option<STInts>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@func")]
    pub func: STFunctionType,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@arg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub arg: Option<STFunctionArgument>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@op")]
    pub op: STFunctionOperator,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    pub value: STFunctionValue,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "alg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg: Option<Box<LayoutAlgorithm>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "presOf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_of: Option<Box<PresentationOf>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "constrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub constr_lst: Option<Box<LayoutConstraints>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "ruleLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule_lst: Option<Box<LayoutRules>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "forEach")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub for_each: Vec<LayoutForEach>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutNode")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout_node: Vec<LayoutNode>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "choose")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub choose: Vec<LayoutChoose>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutOtherwise {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "alg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg: Option<Box<LayoutAlgorithm>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "shape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<DiagramShape>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "presOf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_of: Option<Box<PresentationOf>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "constrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub constr_lst: Option<Box<LayoutConstraints>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "ruleLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule_lst: Option<Box<LayoutRules>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "forEach")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub for_each: Vec<LayoutForEach>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutNode")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout_node: Vec<LayoutNode>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "choose")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub choose: Vec<LayoutChoose>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutChoose {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "if")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub r#if: Vec<LayoutWhen>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "else")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#else: Option<Box<LayoutOtherwise>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct DiagramSampleData {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@useDef")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_def: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "dataModel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_model: Option<Box<DataModel>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct DiagramCategory {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@pri")]
    pub pri: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramCategories {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cat: Vec<DiagramCategory>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramName {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct DiagramDescription {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct DiagramDefinition {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@defStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_style: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<DiagramName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<DiagramDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramCategories>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "sampData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub samp_data: Option<Box<DiagramSampleData>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "styleData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_data: Option<Box<DiagramSampleData>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "clrData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_data: Option<Box<DiagramSampleData>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutNode")]
    pub layout_node: Box<LayoutNode>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmLayoutDef = Box<DiagramDefinition>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramDefinitionHeader {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    pub unique_id: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@defStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_style: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@resId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub res_id: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<DiagramName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<DiagramDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramCategories>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmLayoutDefHdr = Box<DiagramDefinitionHeader>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramDefinitionHeaderList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "layoutDefHdr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub layout_def_hdr: Vec<DiagramDefinitionHeader>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmLayoutDefHdrLst = Box<DiagramDefinitionHeaderList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramRelationshipIds {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@r:dm")]
    pub dm: STRelationshipId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@r:lo")]
    pub lo: STRelationshipId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@r:qs")]
    pub qs: STRelationshipId,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@r:cs")]
    pub cs: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type DdgrmRelIds = Box<DiagramRelationshipIds>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramElementProperties {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presAssocID")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_assoc_i_d: Option<STModelId>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_name: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presStyleLbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_style_lbl: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presStyleIdx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_style_idx: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@presStyleCnt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_style_cnt: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@loTypeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lo_type_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@loCatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lo_cat_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@qsTypeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub qs_type_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@qsCatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub qs_cat_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@csTypeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs_type_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@csCatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs_cat_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@coherent3DOff")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub coherent3_d_off: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@phldrT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phldr_t: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@phldr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub phldr: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_ang: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custFlipVert")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cust_flip_vert: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custFlipHor")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cust_flip_hor: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custSzX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_sz_x: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custSzY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_sz_y: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custScaleX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_scale_x: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custScaleY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_scale_y: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custT")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cust_t: Option<bool>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custLinFactX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_lin_fact_x: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custLinFactY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_lin_fact_y: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custLinFactNeighborX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_lin_fact_neighbor_x: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custLinFactNeighborY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_lin_fact_neighbor_y: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custRadScaleRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_rad_scale_rad: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@custRadScaleInc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_rad_scale_inc: Option<STPrSetCustVal>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "presLayoutVars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pres_layout_vars: Option<Box<LayoutVariableProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct OrgChartProperties {
    #[cfg(feature = "dml-diagrams")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChildMaximum {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STNodeCount>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChildPreference {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STNodeCount>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BulletEnabled {
    #[cfg(feature = "dml-diagrams")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutDirection {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STDirection>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HierarchyBranchStyle {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STHierBranchStyle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AnimateOneByOne {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STAnimOneStr>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AnimateLevel {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STAnimLvlStr>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ResizeHandles {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STResizeHandlesStr>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LayoutVariableProperties {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "orgChart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub org_chart: Option<Box<OrgChartProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "chMax")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ch_max: Option<Box<ChildMaximum>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "chPref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ch_pref: Option<Box<ChildPreference>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "bulletEnabled")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bullet_enabled: Option<Box<BulletEnabled>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<Box<LayoutDirection>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "hierBranch")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hier_branch: Option<Box<HierarchyBranchStyle>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "animOne")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anim_one: Option<Box<AnimateOneByOne>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "animLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anim_lvl: Option<Box<AnimateLevel>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "resizeHandles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub resize_handles: Option<Box<ResizeHandles>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StyleDefinitionName {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct StyleDefinitionDescription {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[cfg(feature = "dml-diagrams")]
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
pub struct DiagramStyleCategory {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@type")]
    pub r#type: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@pri")]
    pub pri: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramStyleCategories {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "cat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cat: Vec<DiagramStyleCategory>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramTextProperties {
    #[cfg(feature = "dml-diagrams")]
    #[serde(skip)]
    #[serde(default)]
    pub text3_d: Option<Box<EGText3D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramStyleLabel {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@name")]
    pub name: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "sp3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp3d: Option<Box<CTShape3D>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "txPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_pr: Option<Box<DiagramTextProperties>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramStyleDefinition {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_id: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<StyleDefinitionName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<StyleDefinitionDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramStyleCategories>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "styleLbl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_lbl: Vec<DiagramStyleLabel>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmStyleDef = Box<DiagramStyleDefinition>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagramStyleDefinitionHeader {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@uniqueId")]
    pub unique_id: String,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@minVer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ver: Option<String>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "@resId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub res_id: Option<i32>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub title: Vec<StyleDefinitionName>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "desc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub desc: Vec<StyleDefinitionDescription>,
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "catLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_lst: Option<Box<DiagramStyleCategories>>,
    #[cfg(feature = "dml-diagrams")]
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
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmStyleDefHdr = Box<DiagramStyleDefinitionHeader>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DiagramStyleDefinitionHeaderList {
    #[cfg(feature = "dml-diagrams")]
    #[serde(rename = "styleDefHdr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_def_hdr: Vec<DiagramStyleDefinitionHeader>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type DdgrmStyleDefHdrLst = Box<DiagramStyleDefinitionHeaderList>;
