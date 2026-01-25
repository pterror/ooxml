//! Serde helpers for OOXML-specific serialization conventions.
//!
//! OOXML uses "1"/"0" for boolean attributes (ECMA-376 Part 1, section 22.9.2.1),
//! not "true"/"false" as Rust's default serde would produce.

/// Serde module for `Option<bool>` fields that serializes as "1"/"0".
///
/// Use with `#[serde(default, with = "ooxml_xml::ooxml_bool", skip_serializing_if = "Option::is_none")]`
pub mod ooxml_bool {
    use serde::{self, Deserialize, Deserializer, Serializer};

    pub fn serialize<S>(value: &Option<bool>, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match value {
            Some(true) => serializer.serialize_str("1"),
            Some(false) => serializer.serialize_str("0"),
            None => serializer.serialize_none(),
        }
    }

    pub fn deserialize<'de, D>(deserializer: D) -> Result<Option<bool>, D::Error>
    where
        D: Deserializer<'de>,
    {
        let opt: Option<String> = Option::deserialize(deserializer)?;
        match opt.as_deref() {
            Some("1") | Some("true") => Ok(Some(true)),
            Some("0") | Some("false") => Ok(Some(false)),
            Some(other) => Err(serde::de::Error::custom(format!(
                "expected boolean value (1/0/true/false), got '{}'",
                other
            ))),
            None => Ok(None),
        }
    }
}

/// Serde module for required `bool` fields that serializes as "1"/"0".
///
/// Use with `#[serde(with = "ooxml_xml::ooxml_bool_required")]`
pub mod ooxml_bool_required {
    use serde::{self, Deserialize, Deserializer, Serializer};

    pub fn serialize<S>(value: &bool, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        if *value {
            serializer.serialize_str("1")
        } else {
            serializer.serialize_str("0")
        }
    }

    pub fn deserialize<'de, D>(deserializer: D) -> Result<bool, D::Error>
    where
        D: Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            other => Err(serde::de::Error::custom(format!(
                "expected boolean value (1/0/true/false), got '{}'",
                other
            ))),
        }
    }
}
