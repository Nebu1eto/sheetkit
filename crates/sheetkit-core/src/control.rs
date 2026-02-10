//! Form control support for Excel worksheets.
//!
//! Excel uses legacy VML (Vector Markup Language) drawing parts for form
//! controls such as buttons, check boxes, option buttons, spin buttons,
//! scroll bars, group boxes, and labels. This module generates the VML
//! markup needed to add form controls and parses existing VML to read
//! them back.

use crate::error::{Error, Result};
use crate::utils::cell_ref::cell_name_to_coordinates;

/// Form control types.
#[derive(Debug, Clone, PartialEq)]
pub enum FormControlType {
    Button,
    CheckBox,
    OptionButton,
    SpinButton,
    ScrollBar,
    GroupBox,
    Label,
}

impl FormControlType {
    /// Return the VML ObjectType string for x:ClientData.
    pub fn object_type(&self) -> &str {
        match self {
            FormControlType::Button => "Button",
            FormControlType::CheckBox => "Checkbox",
            FormControlType::OptionButton => "Radio",
            FormControlType::SpinButton => "Spin",
            FormControlType::ScrollBar => "Scroll",
            FormControlType::GroupBox => "GBox",
            FormControlType::Label => "Label",
        }
    }

    /// Parse an ObjectType string back into a FormControlType.
    pub fn from_object_type(s: &str) -> Option<Self> {
        match s {
            "Button" => Some(FormControlType::Button),
            "Checkbox" => Some(FormControlType::CheckBox),
            "Radio" => Some(FormControlType::OptionButton),
            "Spin" => Some(FormControlType::SpinButton),
            "Scroll" => Some(FormControlType::ScrollBar),
            "GBox" => Some(FormControlType::GroupBox),
            "Label" => Some(FormControlType::Label),
            _ => None,
        }
    }

    /// Parse a user-facing type string into a FormControlType.
    pub fn parse(s: &str) -> Result<Self> {
        match s.to_lowercase().as_str() {
            "button" => Ok(FormControlType::Button),
            "checkbox" | "check_box" | "check" => Ok(FormControlType::CheckBox),
            "optionbutton" | "option_button" | "radio" | "radiobutton" | "radio_button" => {
                Ok(FormControlType::OptionButton)
            }
            "spinbutton" | "spin_button" | "spin" | "spinner" => Ok(FormControlType::SpinButton),
            "scrollbar" | "scroll_bar" | "scroll" => Ok(FormControlType::ScrollBar),
            "groupbox" | "group_box" | "group" => Ok(FormControlType::GroupBox),
            "label" => Ok(FormControlType::Label),
            _ => Err(Error::InvalidArgument(format!(
                "unknown form control type: {s}"
            ))),
        }
    }
}

/// Configuration for adding a form control to a worksheet.
#[derive(Debug, Clone)]
pub struct FormControlConfig {
    /// The type of form control.
    pub control_type: FormControlType,
    /// Anchor cell (top-left corner), e.g. "B2".
    pub cell: String,
    /// Width in points. Uses a sensible default per control type if None.
    pub width: Option<f64>,
    /// Height in points. Uses a sensible default per control type if None.
    pub height: Option<f64>,
    /// Display text (Button, CheckBox, OptionButton, GroupBox, Label).
    pub text: Option<String>,
    /// VBA macro name (Button only).
    pub macro_name: Option<String>,
    /// Linked cell reference for value binding (CheckBox, OptionButton, SpinButton, ScrollBar).
    pub cell_link: Option<String>,
    /// Initial checked state (CheckBox, OptionButton).
    pub checked: Option<bool>,
    /// Minimum value (SpinButton, ScrollBar).
    pub min_value: Option<u32>,
    /// Maximum value (SpinButton, ScrollBar).
    pub max_value: Option<u32>,
    /// Step increment (SpinButton, ScrollBar).
    pub increment: Option<u32>,
    /// Page increment (ScrollBar only).
    pub page_increment: Option<u32>,
    /// Current value (SpinButton, ScrollBar).
    pub current_value: Option<u32>,
    /// Enable 3D shading (default true for most controls).
    pub three_d: Option<bool>,
}

impl FormControlConfig {
    /// Create a Button configuration.
    pub fn button(cell: &str, text: &str) -> Self {
        Self {
            control_type: FormControlType::Button,
            cell: cell.to_string(),
            width: None,
            height: None,
            text: Some(text.to_string()),
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        }
    }

    /// Create a CheckBox configuration.
    pub fn checkbox(cell: &str, text: &str) -> Self {
        Self {
            control_type: FormControlType::CheckBox,
            cell: cell.to_string(),
            width: None,
            height: None,
            text: Some(text.to_string()),
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        }
    }

    /// Create a SpinButton configuration.
    pub fn spin_button(cell: &str, min: u32, max: u32) -> Self {
        Self {
            control_type: FormControlType::SpinButton,
            cell: cell.to_string(),
            width: None,
            height: None,
            text: None,
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: Some(min),
            max_value: Some(max),
            increment: Some(1),
            page_increment: None,
            current_value: Some(min),
            three_d: None,
        }
    }

    /// Create a ScrollBar configuration.
    pub fn scroll_bar(cell: &str, min: u32, max: u32) -> Self {
        Self {
            control_type: FormControlType::ScrollBar,
            cell: cell.to_string(),
            width: None,
            height: None,
            text: None,
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: Some(min),
            max_value: Some(max),
            increment: Some(1),
            page_increment: Some(10),
            current_value: Some(min),
            three_d: None,
        }
    }

    /// Validate the configuration for correctness.
    pub fn validate(&self) -> Result<()> {
        cell_name_to_coordinates(&self.cell)?;

        if let Some(ref cl) = self.cell_link {
            cell_name_to_coordinates(cl)?;
        }

        if let (Some(min), Some(max)) = (self.min_value, self.max_value) {
            if min > max {
                return Err(Error::InvalidArgument(format!(
                    "min_value ({min}) must not exceed max_value ({max})"
                )));
            }
        }

        if let Some(inc) = self.increment {
            if inc == 0 {
                return Err(Error::InvalidArgument(
                    "increment must be greater than 0".to_string(),
                ));
            }
        }

        if let Some(page_inc) = self.page_increment {
            if page_inc == 0 {
                return Err(Error::InvalidArgument(
                    "page_increment must be greater than 0".to_string(),
                ));
            }
        }

        Ok(())
    }
}

/// Information about an existing form control, returned when querying.
#[derive(Debug, Clone, PartialEq)]
pub struct FormControlInfo {
    pub control_type: FormControlType,
    pub cell: String,
    pub text: Option<String>,
    pub macro_name: Option<String>,
    pub cell_link: Option<String>,
    pub checked: Option<bool>,
    pub current_value: Option<u32>,
    pub min_value: Option<u32>,
    pub max_value: Option<u32>,
    pub increment: Option<u32>,
    pub page_increment: Option<u32>,
}

/// Default dimensions (width, height) in points for each control type.
fn default_dimensions(ct: &FormControlType) -> (f64, f64) {
    match ct {
        FormControlType::Button => (72.0, 24.0),
        FormControlType::CheckBox => (72.0, 18.0),
        FormControlType::OptionButton => (72.0, 18.0),
        FormControlType::SpinButton => (15.75, 30.0),
        FormControlType::ScrollBar => (15.75, 60.0),
        FormControlType::GroupBox => (120.0, 72.0),
        FormControlType::Label => (72.0, 18.0),
    }
}

/// VML shapetype id for form controls (differs from comments which use t202).
const FORM_CONTROL_SHAPETYPE_ID: &str = "_x0000_t201";

/// Build the VML anchor string from a cell reference and optional dimensions.
///
/// The anchor format is: col1, col1Off, row1, row1Off, col2, col2Off, row2, row2Off.
/// Offsets are in units of 1/1024 column width or 1/256 row height.
fn build_control_anchor(cell: &str, width_pt: f64, height_pt: f64) -> Result<String> {
    let (col, row) = cell_name_to_coordinates(cell)?;
    let col0 = col - 1;
    let row0 = row - 1;

    // Approximate column and row span from point dimensions.
    // Standard column width is ~64 pixels (~48pt), standard row height ~15pt.
    let col_span = ((width_pt / 48.0).ceil() as u32).max(1);
    let row_span = ((height_pt / 15.0).ceil() as u32).max(1);

    let col2 = col0 + col_span;
    let row2 = row0 + row_span;

    Ok(format!("{col0}, 15, {row0}, 10, {col2}, 63, {row2}, 24"))
}

/// Build a complete VML drawing document containing form control shapes.
///
/// `controls` is a list of FormControlConfig entries that have been validated.
/// `start_shape_id` is the starting shape ID (usually 1025).
/// Returns the VML XML string.
pub fn build_form_control_vml(controls: &[FormControlConfig], start_shape_id: usize) -> String {
    use std::fmt::Write;

    let mut shapes = String::new();
    for (i, config) in controls.iter().enumerate() {
        let shape_id = start_shape_id + i;
        let (default_w, default_h) = default_dimensions(&config.control_type);
        let width = config.width.unwrap_or(default_w);
        let height = config.height.unwrap_or(default_h);

        let anchor = match build_control_anchor(&config.cell, width, height) {
            Ok(a) => a,
            Err(_) => continue,
        };

        write_form_control_shape(&mut shapes, shape_id, i + 1, &anchor, config);
    }

    let mut doc = String::with_capacity(1024 + shapes.len());
    doc.push_str("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
    doc.push_str(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
    doc.push_str(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n");
    doc.push_str(" <o:shapelayout v:ext=\"edit\">\n");
    doc.push_str("  <o:idmap v:ext=\"edit\" data=\"1\"/>\n");
    doc.push_str(" </o:shapelayout>\n");

    // Form control shapetype (t201).
    let _ = write!(
        doc,
        " <v:shapetype id=\"{}\" coordsize=\"21600,21600\" o:spt=\"201\" \
         path=\"m,l,21600r21600,l21600,xe\">\n\
         \x20 <v:stroke joinstyle=\"miter\"/>\n\
         \x20 <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n\
         </v:shapetype>\n",
        FORM_CONTROL_SHAPETYPE_ID,
    );

    doc.push_str(&shapes);
    doc.push_str("</xml>\n");
    doc
}

/// Write a single VML form control shape element.
fn write_form_control_shape(
    out: &mut String,
    shape_id: usize,
    z_index: usize,
    anchor: &str,
    config: &FormControlConfig,
) {
    use std::fmt::Write;

    let _ = write!(out, " <v:shape id=\"_x0000_s{shape_id}\"");
    let _ = write!(out, " type=\"#{FORM_CONTROL_SHAPETYPE_ID}\"");
    let _ = write!(
        out,
        " style=\"position:absolute;z-index:{z_index};visibility:visible\""
    );

    match config.control_type {
        FormControlType::Button => {
            out.push_str(" fillcolor=\"buttonFace\" o:insetmode=\"auto\">\n");
            out.push_str("  <v:fill color2=\"buttonFace\" o:detectmouseclick=\"t\"/>\n");
            out.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
            if let Some(ref text) = config.text {
                let _ = write!(
                    out,
                    "  <v:textbox>\n\
                     \x20  <div style=\"text-align:center\">\
                     <font face=\"Calibri\" size=\"220\" color=\"#000000\">{text}</font>\
                     </div>\n\
                     \x20 </v:textbox>\n"
                );
            }
        }
        FormControlType::CheckBox | FormControlType::OptionButton => {
            out.push_str(" fillcolor=\"window\" o:insetmode=\"auto\">\n");
            out.push_str("  <v:fill color2=\"window\"/>\n");
            if let Some(ref text) = config.text {
                let _ = write!(
                    out,
                    "  <v:textbox>\n\
                     \x20  <div>{text}</div>\n\
                     \x20 </v:textbox>\n"
                );
            }
        }
        FormControlType::SpinButton | FormControlType::ScrollBar => {
            out.push_str(" fillcolor=\"buttonFace\" o:insetmode=\"auto\">\n");
            out.push_str("  <v:fill color2=\"buttonFace\"/>\n");
        }
        FormControlType::GroupBox => {
            out.push_str(" filled=\"f\" stroked=\"f\" o:insetmode=\"auto\">\n");
            if let Some(ref text) = config.text {
                let _ = write!(
                    out,
                    "  <v:textbox>\n\
                     \x20  <div>{text}</div>\n\
                     \x20 </v:textbox>\n"
                );
            }
        }
        FormControlType::Label => {
            out.push_str(" filled=\"f\" stroked=\"f\" o:insetmode=\"auto\">\n");
            if let Some(ref text) = config.text {
                let _ = write!(
                    out,
                    "  <v:textbox>\n\
                     \x20  <div>{text}</div>\n\
                     \x20 </v:textbox>\n"
                );
            }
        }
    }

    // x:ClientData element with control-specific properties.
    let object_type = config.control_type.object_type();
    let _ = writeln!(out, "  <x:ClientData ObjectType=\"{object_type}\">");
    let _ = writeln!(out, "   <x:Anchor>{anchor}</x:Anchor>");
    out.push_str("   <x:PrintObject>False</x:PrintObject>\n");
    out.push_str("   <x:AutoFill>False</x:AutoFill>\n");

    if let Some(ref macro_name) = config.macro_name {
        let _ = writeln!(out, "   <x:FmlaMacro>{macro_name}</x:FmlaMacro>");
    }

    if let Some(ref cell_link) = config.cell_link {
        let _ = writeln!(out, "   <x:FmlaLink>{cell_link}</x:FmlaLink>");
    }

    if let Some(checked) = config.checked {
        let val = if checked { 1 } else { 0 };
        let _ = writeln!(out, "   <x:Checked>{val}</x:Checked>");
    }

    if let Some(val) = config.current_value {
        let _ = writeln!(out, "   <x:Val>{val}</x:Val>");
    }

    if let Some(min) = config.min_value {
        let _ = writeln!(out, "   <x:Min>{min}</x:Min>");
    }

    if let Some(max) = config.max_value {
        let _ = writeln!(out, "   <x:Max>{max}</x:Max>");
    }

    if let Some(inc) = config.increment {
        let _ = writeln!(out, "   <x:Inc>{inc}</x:Inc>");
    }

    if let Some(page_inc) = config.page_increment {
        let _ = writeln!(out, "   <x:Page>{page_inc}</x:Page>");
    }

    // 3D shading: default is true for most controls. Write NoThreeD only when explicitly false.
    let three_d = config.three_d.unwrap_or(true);
    if !three_d {
        out.push_str("   <x:NoThreeD/>\n");
    }

    out.push_str("  </x:ClientData>\n");
    out.push_str(" </v:shape>\n");
}

/// Parse form controls from a VML drawing XML string.
///
/// Scans for `<x:ClientData ObjectType="...">` elements and extracts
/// control properties. Returns a list of FormControlInfo.
pub fn parse_form_controls(vml_xml: &str) -> Vec<FormControlInfo> {
    let mut controls = Vec::new();

    let mut search_from = 0;
    while let Some(shape_start) = vml_xml[search_from..].find("<v:shape ") {
        let abs_start = search_from + shape_start;
        let shape_end = match vml_xml[abs_start..].find("</v:shape>") {
            Some(pos) => abs_start + pos + "</v:shape>".len(),
            None => break,
        };
        let shape_xml = &vml_xml[abs_start..shape_end];

        // Skip non-form-control shapes (e.g. comment shapes use ObjectType="Note").
        if let Some(info) = parse_single_control(shape_xml) {
            controls.push(info);
        }
        search_from = shape_end;
    }

    controls
}

/// Parse a single v:shape element into a FormControlInfo, if it is a form control.
fn parse_single_control(shape_xml: &str) -> Option<FormControlInfo> {
    // Find the ClientData element.
    let cd_start = shape_xml.find("<x:ClientData ")?;
    let cd_end = shape_xml
        .find("</x:ClientData>")
        .map(|p| p + "</x:ClientData>".len())?;
    let cd_xml = &shape_xml[cd_start..cd_end];

    // Extract ObjectType attribute.
    let obj_type = extract_attr(cd_xml, "ObjectType")?;
    let control_type = FormControlType::from_object_type(&obj_type)?;

    // Skip Note types (comments).
    if obj_type == "Note" {
        return None;
    }

    let cell = extract_anchor_cell(cd_xml).unwrap_or_default();
    let text = extract_textbox_text(shape_xml);
    let macro_name = extract_element(cd_xml, "x:FmlaMacro");
    let cell_link = extract_element(cd_xml, "x:FmlaLink");
    let checked = extract_element(cd_xml, "x:Checked").and_then(|v| match v.as_str() {
        "1" => Some(true),
        "0" => Some(false),
        _ => None,
    });
    let current_value = extract_element(cd_xml, "x:Val").and_then(|v| v.parse().ok());
    let min_value = extract_element(cd_xml, "x:Min").and_then(|v| v.parse().ok());
    let max_value = extract_element(cd_xml, "x:Max").and_then(|v| v.parse().ok());
    let increment = extract_element(cd_xml, "x:Inc").and_then(|v| v.parse().ok());
    let page_increment = extract_element(cd_xml, "x:Page").and_then(|v| v.parse().ok());

    Some(FormControlInfo {
        control_type,
        cell,
        text,
        macro_name,
        cell_link,
        checked,
        current_value,
        min_value,
        max_value,
        increment,
        page_increment,
    })
}

/// Extract an XML attribute value from a tag.
fn extract_attr(xml: &str, attr: &str) -> Option<String> {
    let pattern = format!("{attr}=\"");
    let start = xml.find(&pattern)?;
    let val_start = start + pattern.len();
    let end = xml[val_start..].find('"')?;
    Some(xml[val_start..val_start + end].to_string())
}

/// Extract text content of an XML element like `<tag>content</tag>`.
fn extract_element(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let start = xml.find(&open)?;
    let content_start = start + open.len();
    let end = xml[content_start..].find(&close)?;
    let text = xml[content_start..content_start + end].trim().to_string();
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

/// Extract textbox text from a v:shape element.
fn extract_textbox_text(shape_xml: &str) -> Option<String> {
    let tb_start = shape_xml.find("<v:textbox>")?;
    let tb_end = shape_xml.find("</v:textbox>")?;
    let tb_content = &shape_xml[tb_start + "<v:textbox>".len()..tb_end];

    // Text is inside a <div> or <font> element; extract plain text.
    let mut text = String::new();
    let mut in_tag = false;
    for ch in tb_content.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => text.push(ch),
            _ => {}
        }
    }
    let trimmed = text.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

/// Extract the anchor cell reference from x:Anchor element.
///
/// The anchor format is "col1, col1Off, row1, row1Off, col2, col2Off, row2, row2Off".
/// We derive the cell from col1 (0-based) and row1 (0-based).
fn extract_anchor_cell(cd_xml: &str) -> Option<String> {
    let anchor_text = extract_element(cd_xml, "x:Anchor")?;
    let parts: Vec<&str> = anchor_text.split(',').map(|s| s.trim()).collect();
    if parts.len() < 4 {
        return None;
    }
    let col0: u32 = parts[0].parse().ok()?;
    let row0: u32 = parts[2].parse().ok()?;
    crate::utils::cell_ref::coordinates_to_cell_name(col0 + 1, row0 + 1).ok()
}

/// Merge new form control VML into existing VML bytes (for sheets that
/// already have VML content from comments or other controls).
///
/// This appends new shape elements before the closing `</xml>` tag and
/// updates the shapetype if needed.
pub fn merge_vml_controls(
    existing_vml: &[u8],
    controls: &[FormControlConfig],
    start_shape_id: usize,
) -> Vec<u8> {
    let existing_str = String::from_utf8_lossy(existing_vml);

    // Generate shape elements for the new controls.
    let mut shapes = String::new();
    for (i, config) in controls.iter().enumerate() {
        let shape_id = start_shape_id + i;
        let (default_w, default_h) = default_dimensions(&config.control_type);
        let width = config.width.unwrap_or(default_w);
        let height = config.height.unwrap_or(default_h);

        if let Ok(anchor) = build_control_anchor(&config.cell, width, height) {
            write_form_control_shape(&mut shapes, shape_id, shape_id, &anchor, config);
        }
    }

    // Check if the form control shapetype already exists.
    let shapetype_exists = existing_str.contains(FORM_CONTROL_SHAPETYPE_ID);

    let shapetype_xml = if !shapetype_exists {
        format!(
            " <v:shapetype id=\"{FORM_CONTROL_SHAPETYPE_ID}\" coordsize=\"21600,21600\" \
             o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\">\n\
             \x20 <v:stroke joinstyle=\"miter\"/>\n\
             \x20 <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n\
             </v:shapetype>\n"
        )
    } else {
        String::new()
    };

    // Insert before closing </xml>.
    if let Some(close_pos) = existing_str.rfind("</xml>") {
        let mut result =
            String::with_capacity(existing_str.len() + shapetype_xml.len() + shapes.len());
        result.push_str(&existing_str[..close_pos]);
        result.push_str(&shapetype_xml);
        result.push_str(&shapes);
        result.push_str("</xml>\n");
        result.into_bytes()
    } else {
        // Malformed VML; return new VML document.
        build_form_control_vml(controls, start_shape_id).into_bytes()
    }
}

impl FormControlInfo {
    /// Convert a parsed `FormControlInfo` back to a `FormControlConfig`.
    ///
    /// This is used during VML hydration to reconstruct the config list from
    /// existing VML data so that subsequent add/delete operations work correctly.
    pub fn to_config(&self) -> FormControlConfig {
        FormControlConfig {
            control_type: self.control_type.clone(),
            cell: self.cell.clone(),
            width: None,
            height: None,
            text: self.text.clone(),
            macro_name: self.macro_name.clone(),
            cell_link: self.cell_link.clone(),
            checked: self.checked,
            min_value: self.min_value,
            max_value: self.max_value,
            increment: self.increment,
            page_increment: self.page_increment,
            current_value: self.current_value,
            three_d: None,
        }
    }
}

/// Count existing VML shapes in VML bytes to determine the next shape ID.
pub fn count_vml_shapes(vml_bytes: &[u8]) -> usize {
    let vml_str = String::from_utf8_lossy(vml_bytes);
    vml_str.matches("<v:shape ").count()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_form_control_type_parse() {
        assert_eq!(
            FormControlType::parse("button").unwrap(),
            FormControlType::Button
        );
        assert_eq!(
            FormControlType::parse("Button").unwrap(),
            FormControlType::Button
        );
        assert_eq!(
            FormControlType::parse("checkbox").unwrap(),
            FormControlType::CheckBox
        );
        assert_eq!(
            FormControlType::parse("check_box").unwrap(),
            FormControlType::CheckBox
        );
        assert_eq!(
            FormControlType::parse("radio").unwrap(),
            FormControlType::OptionButton
        );
        assert_eq!(
            FormControlType::parse("optionButton").unwrap(),
            FormControlType::OptionButton
        );
        assert_eq!(
            FormControlType::parse("spin").unwrap(),
            FormControlType::SpinButton
        );
        assert_eq!(
            FormControlType::parse("spinner").unwrap(),
            FormControlType::SpinButton
        );
        assert_eq!(
            FormControlType::parse("scroll").unwrap(),
            FormControlType::ScrollBar
        );
        assert_eq!(
            FormControlType::parse("scrollbar").unwrap(),
            FormControlType::ScrollBar
        );
        assert_eq!(
            FormControlType::parse("group").unwrap(),
            FormControlType::GroupBox
        );
        assert_eq!(
            FormControlType::parse("groupbox").unwrap(),
            FormControlType::GroupBox
        );
        assert_eq!(
            FormControlType::parse("label").unwrap(),
            FormControlType::Label
        );
        assert!(FormControlType::parse("unknown").is_err());
    }

    #[test]
    fn test_form_control_type_object_type() {
        assert_eq!(FormControlType::Button.object_type(), "Button");
        assert_eq!(FormControlType::CheckBox.object_type(), "Checkbox");
        assert_eq!(FormControlType::OptionButton.object_type(), "Radio");
        assert_eq!(FormControlType::SpinButton.object_type(), "Spin");
        assert_eq!(FormControlType::ScrollBar.object_type(), "Scroll");
        assert_eq!(FormControlType::GroupBox.object_type(), "GBox");
        assert_eq!(FormControlType::Label.object_type(), "Label");
    }

    #[test]
    fn test_form_control_type_roundtrip() {
        let types = vec![
            FormControlType::Button,
            FormControlType::CheckBox,
            FormControlType::OptionButton,
            FormControlType::SpinButton,
            FormControlType::ScrollBar,
            FormControlType::GroupBox,
            FormControlType::Label,
        ];
        for ct in types {
            let obj_type = ct.object_type();
            let parsed = FormControlType::from_object_type(obj_type).unwrap();
            assert_eq!(parsed, ct);
        }
    }

    #[test]
    fn test_validate_config_valid() {
        let config = FormControlConfig::button("A1", "Click Me");
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_validate_config_invalid_cell() {
        let mut config = FormControlConfig::button("INVALID", "Click Me");
        config.cell = "ZZZZZ".to_string();
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_validate_config_min_exceeds_max() {
        let mut config = FormControlConfig::spin_button("A1", 0, 100);
        config.min_value = Some(200);
        config.max_value = Some(100);
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_validate_config_zero_increment() {
        let mut config = FormControlConfig::spin_button("A1", 0, 100);
        config.increment = Some(0);
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_validate_config_zero_page_increment() {
        let mut config = FormControlConfig::scroll_bar("A1", 0, 100);
        config.page_increment = Some(0);
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_validate_config_invalid_cell_link() {
        let mut config = FormControlConfig::checkbox("A1", "Check");
        config.cell_link = Some("NOT_A_CELL".to_string());
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_build_button_vml() {
        let config = FormControlConfig::button("B2", "Click Me");
        let vml = build_form_control_vml(&[config], 1025);

        assert!(vml.contains("xmlns:v=\"urn:schemas-microsoft-com:vml\""));
        assert!(vml.contains("xmlns:o=\"urn:schemas-microsoft-com:office:office\""));
        assert!(vml.contains("xmlns:x=\"urn:schemas-microsoft-com:office:excel\""));
        assert!(vml.contains("ObjectType=\"Button\""));
        assert!(vml.contains("Click Me"));
        assert!(vml.contains("_x0000_s1025"));
        assert!(vml.contains("_x0000_t201"));
        assert!(vml.contains("fillcolor=\"buttonFace\""));
    }

    #[test]
    fn test_build_checkbox_vml() {
        let mut config = FormControlConfig::checkbox("A1", "Enable Feature");
        config.cell_link = Some("$C$1".to_string());
        config.checked = Some(true);

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("ObjectType=\"Checkbox\""));
        assert!(vml.contains("Enable Feature"));
        assert!(vml.contains("<x:FmlaLink>$C$1</x:FmlaLink>"));
        assert!(vml.contains("<x:Checked>1</x:Checked>"));
    }

    #[test]
    fn test_build_option_button_vml() {
        let config = FormControlConfig {
            control_type: FormControlType::OptionButton,
            cell: "A3".to_string(),
            width: None,
            height: None,
            text: Some("Option A".to_string()),
            macro_name: None,
            cell_link: Some("$D$1".to_string()),
            checked: Some(false),
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        };

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("ObjectType=\"Radio\""));
        assert!(vml.contains("Option A"));
        assert!(vml.contains("<x:FmlaLink>$D$1</x:FmlaLink>"));
        assert!(vml.contains("<x:Checked>0</x:Checked>"));
    }

    #[test]
    fn test_build_spin_button_vml() {
        let config = FormControlConfig::spin_button("E1", 0, 100);
        let vml = build_form_control_vml(&[config], 1025);

        assert!(vml.contains("ObjectType=\"Spin\""));
        assert!(vml.contains("<x:Min>0</x:Min>"));
        assert!(vml.contains("<x:Max>100</x:Max>"));
        assert!(vml.contains("<x:Inc>1</x:Inc>"));
        assert!(vml.contains("<x:Val>0</x:Val>"));
    }

    #[test]
    fn test_build_scroll_bar_vml() {
        let config = FormControlConfig::scroll_bar("F1", 10, 200);
        let vml = build_form_control_vml(&[config], 1025);

        assert!(vml.contains("ObjectType=\"Scroll\""));
        assert!(vml.contains("<x:Min>10</x:Min>"));
        assert!(vml.contains("<x:Max>200</x:Max>"));
        assert!(vml.contains("<x:Inc>1</x:Inc>"));
        assert!(vml.contains("<x:Page>10</x:Page>"));
    }

    #[test]
    fn test_build_group_box_vml() {
        let config = FormControlConfig {
            control_type: FormControlType::GroupBox,
            cell: "A1".to_string(),
            width: None,
            height: None,
            text: Some("Options".to_string()),
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        };

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("ObjectType=\"GBox\""));
        assert!(vml.contains("Options"));
        assert!(vml.contains("filled=\"f\""));
    }

    #[test]
    fn test_build_label_vml() {
        let config = FormControlConfig {
            control_type: FormControlType::Label,
            cell: "A1".to_string(),
            width: None,
            height: None,
            text: Some("Status:".to_string()),
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        };

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("ObjectType=\"Label\""));
        assert!(vml.contains("Status:"));
    }

    #[test]
    fn test_build_button_with_macro() {
        let mut config = FormControlConfig::button("A1", "Run Macro");
        config.macro_name = Some("Sheet1.MyMacro".to_string());

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("<x:FmlaMacro>Sheet1.MyMacro</x:FmlaMacro>"));
    }

    #[test]
    fn test_build_control_no_three_d() {
        let mut config = FormControlConfig::checkbox("A1", "Flat");
        config.three_d = Some(false);

        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("<x:NoThreeD/>"));
    }

    #[test]
    fn test_build_multiple_controls() {
        let controls = vec![
            FormControlConfig::button("A1", "Button 1"),
            FormControlConfig::checkbox("A3", "Check 1"),
            FormControlConfig::spin_button("C1", 0, 50),
        ];

        let vml = build_form_control_vml(&controls, 1025);
        assert!(vml.contains("_x0000_s1025"));
        assert!(vml.contains("_x0000_s1026"));
        assert!(vml.contains("_x0000_s1027"));
        assert!(vml.contains("ObjectType=\"Button\""));
        assert!(vml.contains("ObjectType=\"Checkbox\""));
        assert!(vml.contains("ObjectType=\"Spin\""));
    }

    #[test]
    fn test_parse_form_controls_button() {
        let config = FormControlConfig::button("B2", "Click Me");
        let vml = build_form_control_vml(&[config], 1025);

        let controls = parse_form_controls(&vml);
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[0].text.as_deref(), Some("Click Me"));
    }

    #[test]
    fn test_parse_form_controls_checkbox_with_link() {
        let mut config = FormControlConfig::checkbox("A1", "Toggle");
        config.cell_link = Some("$D$1".to_string());
        config.checked = Some(true);
        let vml = build_form_control_vml(&[config], 1025);

        let controls = parse_form_controls(&vml);
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::CheckBox);
        assert_eq!(controls[0].text.as_deref(), Some("Toggle"));
        assert_eq!(controls[0].cell_link.as_deref(), Some("$D$1"));
        assert_eq!(controls[0].checked, Some(true));
    }

    #[test]
    fn test_parse_form_controls_spin_button() {
        let config = FormControlConfig::spin_button("C1", 5, 50);
        let vml = build_form_control_vml(&[config], 1025);

        let controls = parse_form_controls(&vml);
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::SpinButton);
        assert_eq!(controls[0].min_value, Some(5));
        assert_eq!(controls[0].max_value, Some(50));
        assert_eq!(controls[0].increment, Some(1));
        assert_eq!(controls[0].current_value, Some(5));
    }

    #[test]
    fn test_parse_form_controls_scroll_bar() {
        let config = FormControlConfig::scroll_bar("E1", 0, 100);
        let vml = build_form_control_vml(&[config], 1025);

        let controls = parse_form_controls(&vml);
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::ScrollBar);
        assert_eq!(controls[0].page_increment, Some(10));
    }

    #[test]
    fn test_parse_multiple_controls() {
        let controls = vec![
            FormControlConfig::button("A1", "Btn"),
            FormControlConfig::checkbox("A3", "Chk"),
            FormControlConfig::spin_button("C1", 0, 10),
        ];
        let vml = build_form_control_vml(&controls, 1025);

        let parsed = parse_form_controls(&vml);
        assert_eq!(parsed.len(), 3);
        assert_eq!(parsed[0].control_type, FormControlType::Button);
        assert_eq!(parsed[1].control_type, FormControlType::CheckBox);
        assert_eq!(parsed[2].control_type, FormControlType::SpinButton);
    }

    #[test]
    fn test_parse_ignores_comment_shapes() {
        // Build a VML that has both a comment (Note) and a form control.
        let comment_vml = crate::vml::build_vml_drawing(&["A1"]);
        let controls = parse_form_controls(&comment_vml);
        assert!(controls.is_empty(), "comment shapes should be ignored");
    }

    #[test]
    fn test_count_vml_shapes() {
        let vml = build_form_control_vml(
            &[
                FormControlConfig::button("A1", "B1"),
                FormControlConfig::checkbox("A3", "C1"),
            ],
            1025,
        );
        assert_eq!(count_vml_shapes(vml.as_bytes()), 2);
    }

    #[test]
    fn test_merge_vml_controls() {
        let existing = build_form_control_vml(&[FormControlConfig::button("A1", "First")], 1025);
        let new_controls = vec![FormControlConfig::checkbox("A3", "Second")];
        let merged = merge_vml_controls(existing.as_bytes(), &new_controls, 1026);
        let merged_str = String::from_utf8(merged).unwrap();

        assert!(merged_str.contains("_x0000_s1025"));
        assert!(merged_str.contains("_x0000_s1026"));
        assert!(merged_str.contains("ObjectType=\"Button\""));
        assert!(merged_str.contains("ObjectType=\"Checkbox\""));
        // Should not duplicate shapetype.
        let shapetype_count = merged_str.matches(FORM_CONTROL_SHAPETYPE_ID).count();
        // One in the shapetype definition and references in each shape = 1 (def) + 2 (refs)
        assert!(shapetype_count >= 2);
    }

    #[test]
    fn test_build_control_anchor_basic() {
        let anchor = build_control_anchor("A1", 72.0, 24.0).unwrap();
        let parts: Vec<&str> = anchor.split(", ").collect();
        assert_eq!(parts.len(), 8);
        assert_eq!(parts[0], "0"); // col0
        assert_eq!(parts[2], "0"); // row0
    }

    #[test]
    fn test_build_control_anchor_offset_cell() {
        let anchor = build_control_anchor("C5", 72.0, 30.0).unwrap();
        let parts: Vec<&str> = anchor.split(", ").collect();
        assert_eq!(parts[0], "2"); // col0 (C = 3, 0-based = 2)
        assert_eq!(parts[2], "4"); // row0 (5, 0-based = 4)
    }

    #[test]
    fn test_build_control_anchor_invalid_cell() {
        assert!(build_control_anchor("INVALID", 72.0, 24.0).is_err());
    }

    #[test]
    fn test_default_dimensions() {
        let (w, h) = default_dimensions(&FormControlType::Button);
        assert_eq!(w, 72.0);
        assert_eq!(h, 24.0);

        let (w, h) = default_dimensions(&FormControlType::ScrollBar);
        assert_eq!(w, 15.75);
        assert_eq!(h, 60.0);
    }

    #[test]
    fn test_extract_anchor_cell() {
        let cd = "<x:ClientData ObjectType=\"Button\"><x:Anchor>1, 15, 0, 10, 3, 63, 2, 24</x:Anchor></x:ClientData>";
        let cell = extract_anchor_cell(cd).unwrap();
        assert_eq!(cell, "B1");
    }

    #[test]
    fn test_custom_dimensions() {
        let mut config = FormControlConfig::button("A1", "Wide");
        config.width = Some(200.0);
        config.height = Some(50.0);
        let vml = build_form_control_vml(&[config], 1025);
        assert!(vml.contains("_x0000_s1025"));
    }

    #[test]
    fn test_workbook_add_form_control() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        let config = FormControlConfig::button("B2", "Click Me");
        wb.add_form_control("Sheet1", config).unwrap();

        let controls = wb.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[0].text.as_deref(), Some("Click Me"));
    }

    #[test]
    fn test_workbook_add_multiple_form_controls() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Button 1"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Check 1"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::spin_button("C1", 0, 100))
            .unwrap();

        let controls = wb.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 3);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[1].control_type, FormControlType::CheckBox);
        assert_eq!(controls[2].control_type, FormControlType::SpinButton);
    }

    #[test]
    fn test_workbook_add_form_control_sheet_not_found() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        let config = FormControlConfig::button("A1", "Test");
        let result = wb.add_form_control("NoSheet", config);
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_add_form_control_invalid_cell() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        let config = FormControlConfig {
            control_type: FormControlType::Button,
            cell: "INVALID".to_string(),
            width: None,
            height: None,
            text: Some("Test".to_string()),
            macro_name: None,
            cell_link: None,
            checked: None,
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        };
        let result = wb.add_form_control("Sheet1", config);
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_delete_form_control() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Btn 1"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Chk 1"))
            .unwrap();

        wb.delete_form_control("Sheet1", 0).unwrap();

        let controls = wb.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::CheckBox);
    }

    #[test]
    fn test_workbook_delete_form_control_out_of_bounds() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Btn"))
            .unwrap();
        let result = wb.delete_form_control("Sheet1", 5);
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_form_control_save_roundtrip() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("form_controls.xlsx");

        let mut wb = Workbook::new();
        let mut btn = FormControlConfig::button("B2", "Submit");
        btn.macro_name = Some("Sheet1.OnSubmit".to_string());
        wb.add_form_control("Sheet1", btn).unwrap();

        let mut chk = FormControlConfig::checkbox("B4", "Agree");
        chk.cell_link = Some("$D$4".to_string());
        chk.checked = Some(true);
        wb.add_form_control("Sheet1", chk).unwrap();

        wb.add_form_control("Sheet1", FormControlConfig::spin_button("E2", 0, 100))
            .unwrap();

        wb.save(&path).unwrap();

        // Verify VML part exists in the ZIP.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        let has_vml = (1..=10).any(|i| {
            archive
                .by_name(&format!("xl/drawings/vmlDrawing{i}.vml"))
                .is_ok()
        });
        assert!(has_vml, "should have a vmlDrawing file in the ZIP");

        // Re-open and verify controls are preserved.
        let mut wb2 = Workbook::open(&path).unwrap();
        let controls = wb2.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 3);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[0].text.as_deref(), Some("Submit"));
        assert_eq!(controls[0].macro_name.as_deref(), Some("Sheet1.OnSubmit"));
        assert_eq!(controls[1].control_type, FormControlType::CheckBox);
        assert_eq!(controls[1].cell_link.as_deref(), Some("$D$4"));
        assert_eq!(controls[1].checked, Some(true));
        assert_eq!(controls[2].control_type, FormControlType::SpinButton);
        assert_eq!(controls[2].min_value, Some(0));
        assert_eq!(controls[2].max_value, Some(100));
    }

    #[test]
    fn test_workbook_form_control_all_7_types() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Button"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Checkbox"))
            .unwrap();
        wb.add_form_control(
            "Sheet1",
            FormControlConfig {
                control_type: FormControlType::OptionButton,
                cell: "A5".to_string(),
                width: None,
                height: None,
                text: Some("Option".to_string()),
                macro_name: None,
                cell_link: None,
                checked: None,
                min_value: None,
                max_value: None,
                increment: None,
                page_increment: None,
                current_value: None,
                three_d: None,
            },
        )
        .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::spin_button("C1", 0, 10))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::scroll_bar("E1", 0, 100))
            .unwrap();
        wb.add_form_control(
            "Sheet1",
            FormControlConfig {
                control_type: FormControlType::GroupBox,
                cell: "G1".to_string(),
                width: None,
                height: None,
                text: Some("Group".to_string()),
                macro_name: None,
                cell_link: None,
                checked: None,
                min_value: None,
                max_value: None,
                increment: None,
                page_increment: None,
                current_value: None,
                three_d: None,
            },
        )
        .unwrap();
        wb.add_form_control(
            "Sheet1",
            FormControlConfig {
                control_type: FormControlType::Label,
                cell: "I1".to_string(),
                width: None,
                height: None,
                text: Some("Label Text".to_string()),
                macro_name: None,
                cell_link: None,
                checked: None,
                min_value: None,
                max_value: None,
                increment: None,
                page_increment: None,
                current_value: None,
                three_d: None,
            },
        )
        .unwrap();

        let controls = wb.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 7);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[1].control_type, FormControlType::CheckBox);
        assert_eq!(controls[2].control_type, FormControlType::OptionButton);
        assert_eq!(controls[3].control_type, FormControlType::SpinButton);
        assert_eq!(controls[4].control_type, FormControlType::ScrollBar);
        assert_eq!(controls[5].control_type, FormControlType::GroupBox);
        assert_eq!(controls[6].control_type, FormControlType::Label);
    }

    #[test]
    fn test_workbook_form_control_with_comments() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("controls_and_comments.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "A comment".to_string(),
            },
        )
        .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::button("C1", "Button"))
            .unwrap();
        wb.save(&path).unwrap();

        let mut wb2 = Workbook::open(&path).unwrap();
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        let controls = wb2.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].control_type, FormControlType::Button);
    }

    #[test]
    fn test_workbook_get_form_controls_empty() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        let controls = wb.get_form_controls("Sheet1").unwrap();
        assert!(controls.is_empty());
    }

    #[test]
    fn test_workbook_get_form_controls_sheet_not_found() {
        use crate::workbook::Workbook;

        let mut wb = Workbook::new();
        let result = wb.get_form_controls("NoSheet");
        assert!(result.is_err());
    }

    #[test]
    fn test_open_file_get_form_controls_returns_existing() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("get_existing.xlsx");

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Existing"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Check"))
            .unwrap();
        wb.save(&path).unwrap();

        let mut wb2 = Workbook::open(&path).unwrap();
        let controls = wb2.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[0].text.as_deref(), Some("Existing"));
        assert_eq!(controls[1].control_type, FormControlType::CheckBox);
        assert_eq!(controls[1].text.as_deref(), Some("Check"));
    }

    #[test]
    fn test_open_file_add_form_control_preserves_existing() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("add_preserves.xlsx");

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "First"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Second"))
            .unwrap();
        wb.save(&path).unwrap();

        let mut wb2 = Workbook::open(&path).unwrap();
        wb2.add_form_control("Sheet1", FormControlConfig::spin_button("C1", 0, 50))
            .unwrap();

        let controls = wb2.get_form_controls("Sheet1").unwrap();
        assert_eq!(
            controls.len(),
            3,
            "old + new controls should all be present"
        );
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[0].text.as_deref(), Some("First"));
        assert_eq!(controls[1].control_type, FormControlType::CheckBox);
        assert_eq!(controls[1].text.as_deref(), Some("Second"));
        assert_eq!(controls[2].control_type, FormControlType::SpinButton);
        assert_eq!(controls[2].min_value, Some(0));
        assert_eq!(controls[2].max_value, Some(50));
    }

    #[test]
    fn test_open_file_delete_form_control_works() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("delete_works.xlsx");

        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "First"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Second"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::spin_button("C1", 0, 10))
            .unwrap();
        wb.save(&path).unwrap();

        let mut wb2 = Workbook::open(&path).unwrap();
        wb2.delete_form_control("Sheet1", 1).unwrap();

        let controls = wb2.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].control_type, FormControlType::Button);
        assert_eq!(controls[1].control_type, FormControlType::SpinButton);
    }

    #[test]
    fn test_open_file_modify_save_reopen_persistence() {
        use crate::workbook::Workbook;
        use tempfile::TempDir;

        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("persistence_step1.xlsx");
        let path2 = dir.path().join("persistence_step2.xlsx");

        // Step 1: Create with 2 controls.
        let mut wb = Workbook::new();
        wb.add_form_control("Sheet1", FormControlConfig::button("A1", "Button"))
            .unwrap();
        wb.add_form_control("Sheet1", FormControlConfig::checkbox("A3", "Check"))
            .unwrap();
        wb.save(&path1).unwrap();

        // Step 2: Open, add one, delete one, save.
        let mut wb2 = Workbook::open(&path1).unwrap();
        wb2.add_form_control("Sheet1", FormControlConfig::scroll_bar("E1", 0, 100))
            .unwrap();
        wb2.delete_form_control("Sheet1", 0).unwrap();
        wb2.save(&path2).unwrap();

        // Step 3: Re-open and verify.
        let mut wb3 = Workbook::open(&path2).unwrap();
        let controls = wb3.get_form_controls("Sheet1").unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].control_type, FormControlType::CheckBox);
        assert_eq!(controls[0].text.as_deref(), Some("Check"));
        assert_eq!(controls[1].control_type, FormControlType::ScrollBar);
        assert_eq!(controls[1].min_value, Some(0));
        assert_eq!(controls[1].max_value, Some(100));
    }

    #[test]
    fn test_info_to_config_roundtrip() {
        let config = FormControlConfig {
            control_type: FormControlType::CheckBox,
            cell: "B2".to_string(),
            width: None,
            height: None,
            text: Some("Toggle".to_string()),
            macro_name: Some("MyMacro".to_string()),
            cell_link: Some("$D$1".to_string()),
            checked: Some(true),
            min_value: None,
            max_value: None,
            increment: None,
            page_increment: None,
            current_value: None,
            three_d: None,
        };

        let vml = build_form_control_vml(&[config.clone()], 1025);
        let parsed = parse_form_controls(&vml);
        assert_eq!(parsed.len(), 1);
        let roundtripped = parsed[0].to_config();
        assert_eq!(roundtripped.control_type, config.control_type);
        assert_eq!(roundtripped.text, config.text);
        assert_eq!(roundtripped.macro_name, config.macro_name);
        assert_eq!(roundtripped.cell_link, config.cell_link);
        assert_eq!(roundtripped.checked, config.checked);
    }
}
