//! Office Math Markup Language (OMML) support for the ooxml library.
//!
//! This crate provides types and parsing for mathematical formulas
//! in Word, Excel, and PowerPoint documents.
//!
//! # Example
//!
//! ```
//! use ooxml_omml::{MathZone, parse_math_zone};
//!
//! let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
//!     <m:f>
//!         <m:num><m:r><m:t>1</m:t></m:r></m:num>
//!         <m:den><m:r><m:t>2</m:t></m:r></m:den>
//!     </m:f>
//! </m:oMath>"#;
//!
//! let zone = parse_math_zone(xml.as_bytes()).unwrap();
//! assert_eq!(zone.text(), "(1)/(2)");
//! ```
//!
//! # Supported Elements
//!
//! - Fractions (`m:f`)
//! - Radicals (`m:rad`)
//! - N-ary operators like summation and integrals (`m:nary`)
//! - Subscripts and superscripts (`m:sSub`, `m:sSup`, `m:sSubSup`)
//! - Delimiters/parentheses (`m:d`)
//! - Matrices (`m:m`)
//! - Functions (`m:func`)
//! - Accents (`m:acc`)
//! - And more...

pub mod error;
pub mod math;

pub use error::{Error, Result};
pub use math::{
    Accent, Bar, BorderBox, Delimiter, EquationArray, Fraction, FractionType, Function, GroupChar,
    Limit, LimitLocation, MathBox, MathElement, MathRun, MathRunProperties, MathScript, MathStyle,
    MathZone, Matrix, Nary, Phantom, PreScript, Radical, Script, SubSuperscript, VerticalPosition,
    parse_math_zone, parse_math_zone_from_reader,
};
