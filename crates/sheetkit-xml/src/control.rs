//! Worksheet form-control XML schema structures.

use serde::{Deserialize, Serialize};

/// Worksheet controls container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Controls {
    #[serde(rename = "control", default)]
    pub controls: Vec<Control>,
}

/// A worksheet control linked to a control-property part.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Control {
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
    #[serde(rename = "@name", skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}
