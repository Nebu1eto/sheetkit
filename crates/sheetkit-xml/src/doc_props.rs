//! Document properties XML schema structures.
//!
//! Covers:
//! - Core properties (`docProps/core.xml`) - Dublin Core metadata
//! - Extended properties (`docProps/app.xml`) - application metadata
//! - Custom properties (`docProps/custom.xml`) - user-defined key/value pairs

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader;
use quick_xml::Writer;
use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Core document properties (docProps/core.xml).
///
/// Uses Dublin Core namespaces (`dc:`, `dcterms:`, `cp:`).
/// Because quick-xml serde does not handle namespace prefixes well,
/// serialization and deserialization are done manually.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CoreProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
    pub category: Option<String>,
    pub content_status: Option<String>,
}

/// Serialize `CoreProperties` to its XML string representation.
pub fn serialize_core_properties(props: &CoreProperties) -> String {
    let mut writer = Writer::new(Vec::new());

    // XML declaration
    writer
        .write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))
        .unwrap();

    // Root element with namespaces
    let mut root = BytesStart::new("cp:coreProperties");
    root.push_attribute(("xmlns:cp", namespaces::CORE_PROPERTIES));
    root.push_attribute(("xmlns:dc", namespaces::DC));
    root.push_attribute(("xmlns:dcterms", namespaces::DC_TERMS));
    root.push_attribute(("xmlns:dcmitype", DC_MITYPE));
    root.push_attribute(("xmlns:xsi", namespaces::XSI));
    writer.write_event(Event::Start(root)).unwrap();

    // Helper: write a simple text element
    fn write_element(writer: &mut Writer<Vec<u8>>, tag: &str, value: &str) {
        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .unwrap();
        writer
            .write_event(Event::Text(BytesText::new(value)))
            .unwrap();
        writer.write_event(Event::End(BytesEnd::new(tag))).unwrap();
    }

    // Helper: write dcterms element with xsi:type attribute
    fn write_dcterms_element(writer: &mut Writer<Vec<u8>>, tag: &str, value: &str) {
        let mut start = BytesStart::new(tag);
        start.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        writer.write_event(Event::Start(start)).unwrap();
        writer
            .write_event(Event::Text(BytesText::new(value)))
            .unwrap();
        writer.write_event(Event::End(BytesEnd::new(tag))).unwrap();
    }

    if let Some(ref v) = props.title {
        write_element(&mut writer, "dc:title", v);
    }
    if let Some(ref v) = props.subject {
        write_element(&mut writer, "dc:subject", v);
    }
    if let Some(ref v) = props.creator {
        write_element(&mut writer, "dc:creator", v);
    }
    if let Some(ref v) = props.keywords {
        write_element(&mut writer, "cp:keywords", v);
    }
    if let Some(ref v) = props.description {
        write_element(&mut writer, "dc:description", v);
    }
    if let Some(ref v) = props.last_modified_by {
        write_element(&mut writer, "cp:lastModifiedBy", v);
    }
    if let Some(ref v) = props.revision {
        write_element(&mut writer, "cp:revision", v);
    }
    if let Some(ref v) = props.created {
        write_dcterms_element(&mut writer, "dcterms:created", v);
    }
    if let Some(ref v) = props.modified {
        write_dcterms_element(&mut writer, "dcterms:modified", v);
    }
    if let Some(ref v) = props.category {
        write_element(&mut writer, "cp:category", v);
    }
    if let Some(ref v) = props.content_status {
        write_element(&mut writer, "cp:contentStatus", v);
    }

    writer
        .write_event(Event::End(BytesEnd::new("cp:coreProperties")))
        .unwrap();

    String::from_utf8(writer.into_inner()).unwrap()
}

fn xml_error(e: impl std::fmt::Display) -> String {
    format!("XML parse error: {e}")
}

/// Append the decoded content of a text or CDATA event to `buf`.
///
/// Other events are ignored. Entity references in text are unescaped;
/// CDATA content is taken verbatim.
fn append_text(buf: &mut String, event: &Event<'_>) -> Result<(), String> {
    match event {
        Event::Text(t) => buf.push_str(&t.unescape().map_err(xml_error)?),
        Event::CData(c) => buf.push_str(std::str::from_utf8(c).map_err(xml_error)?),
        _ => {}
    }
    Ok(())
}

/// Store `text` into the core property field named by the local name `tag`.
fn set_core_field(props: &mut CoreProperties, tag: &[u8], text: String) {
    let slot = match tag {
        b"title" => &mut props.title,
        b"subject" => &mut props.subject,
        b"creator" => &mut props.creator,
        b"keywords" => &mut props.keywords,
        b"description" => &mut props.description,
        b"lastModifiedBy" => &mut props.last_modified_by,
        b"revision" => &mut props.revision,
        b"created" => &mut props.created,
        b"modified" => &mut props.modified,
        b"category" => &mut props.category,
        b"contentStatus" => &mut props.content_status,
        _ => return,
    };
    *slot = Some(text);
}

/// Deserialize `CoreProperties` from an XML string.
///
/// Fields are read from direct children of the root element, matched by local
/// name regardless of namespace prefix. Values keep their exact text,
/// including empty and whitespace-only strings and CDATA sections.
pub fn deserialize_core_properties(xml: &str) -> Result<CoreProperties, String> {
    let mut reader = Reader::from_str(xml);

    let mut props = CoreProperties::default();
    let mut depth = 0usize;
    // Local name and accumulated text of the field element currently open.
    let mut current: Option<(Vec<u8>, String)> = None;

    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match &event {
            Event::Start(e) => {
                depth += 1;
                if depth == 2 {
                    current = Some((e.local_name().as_ref().to_vec(), String::new()));
                }
            }
            Event::Empty(e) => {
                if depth == 1 {
                    set_core_field(&mut props, e.local_name().as_ref(), String::new());
                }
            }
            Event::Text(_) | Event::CData(_) => {
                if depth == 2 {
                    if let Some((_, buf)) = current.as_mut() {
                        append_text(buf, &event)?;
                    }
                }
            }
            Event::End(_) => {
                if depth == 2 {
                    if let Some((tag, text)) = current.take() {
                        set_core_field(&mut props, &tag, text);
                    }
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

// DCMI Type namespace (not in namespaces.rs because it's only used here)
const DC_MITYPE: &str = "http://purl.org/dc/dcmitype/";

/// Extended (application) properties (`docProps/app.xml`).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename = "Properties")]
pub struct ExtendedProperties {
    #[serde(rename = "@xmlns", default = "default_extended_properties_ns")]
    pub xmlns: String,
    #[serde(rename = "@xmlns:vt", skip_serializing_if = "Option::is_none")]
    pub xmlns_vt: Option<String>,

    #[serde(rename = "Application", skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,
    #[serde(rename = "DocSecurity", skip_serializing_if = "Option::is_none")]
    pub doc_security: Option<u32>,
    #[serde(rename = "ScaleCrop", skip_serializing_if = "Option::is_none")]
    pub scale_crop: Option<bool>,
    #[serde(rename = "Company", skip_serializing_if = "Option::is_none")]
    pub company: Option<String>,
    #[serde(rename = "LinksUpToDate", skip_serializing_if = "Option::is_none")]
    pub links_up_to_date: Option<bool>,
    #[serde(rename = "SharedDoc", skip_serializing_if = "Option::is_none")]
    pub shared_doc: Option<bool>,
    #[serde(rename = "HyperlinksChanged", skip_serializing_if = "Option::is_none")]
    pub hyperlinks_changed: Option<bool>,
    #[serde(rename = "AppVersion", skip_serializing_if = "Option::is_none")]
    pub app_version: Option<String>,
    #[serde(rename = "Template", skip_serializing_if = "Option::is_none")]
    pub template: Option<String>,
    #[serde(rename = "Manager", skip_serializing_if = "Option::is_none")]
    pub manager: Option<String>,
}

impl ExtendedProperties {
    /// Create a new `ExtendedProperties` with the standard namespace set.
    pub fn with_defaults() -> Self {
        Self {
            xmlns: namespaces::EXTENDED_PROPERTIES.to_string(),
            xmlns_vt: Some(namespaces::VT.to_string()),
            ..Default::default()
        }
    }
}

/// Custom properties collection (`docProps/custom.xml`).
///
/// Because the child value elements use a `vt:` namespace prefix that
/// quick-xml serde cannot handle, serialization/deserialization is manual.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CustomProperties {
    pub properties: Vec<CustomProperty>,
}

/// A single custom property entry.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomProperty {
    pub fmtid: String,
    pub pid: u32,
    pub name: String,
    pub value: CustomPropertyValue,
}

/// The typed value of a custom property.
#[derive(Debug, Clone, PartialEq)]
pub enum CustomPropertyValue {
    String(String),
    Int(i32),
    Float(f64),
    Bool(bool),
    DateTime(String),
}

/// Standard fmtid used for custom properties.
pub const CUSTOM_PROPERTY_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/// Serialize `CustomProperties` to its XML string representation.
pub fn serialize_custom_properties(props: &CustomProperties) -> String {
    let mut writer = Writer::new(Vec::new());

    // XML declaration
    writer
        .write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))
        .unwrap();

    let mut root = BytesStart::new("Properties");
    root.push_attribute(("xmlns", namespaces::CUSTOM_PROPERTIES));
    root.push_attribute(("xmlns:vt", namespaces::VT));
    writer.write_event(Event::Start(root)).unwrap();

    for prop in &props.properties {
        let mut elem = BytesStart::new("property");
        elem.push_attribute(("fmtid", prop.fmtid.as_str()));
        elem.push_attribute(("pid", prop.pid.to_string().as_str()));
        elem.push_attribute(("name", prop.name.as_str()));
        writer.write_event(Event::Start(elem)).unwrap();

        match &prop.value {
            CustomPropertyValue::String(s) => {
                writer
                    .write_event(Event::Start(BytesStart::new("vt:lpwstr")))
                    .unwrap();
                writer.write_event(Event::Text(BytesText::new(s))).unwrap();
                writer
                    .write_event(Event::End(BytesEnd::new("vt:lpwstr")))
                    .unwrap();
            }
            CustomPropertyValue::Int(n) => {
                writer
                    .write_event(Event::Start(BytesStart::new("vt:i4")))
                    .unwrap();
                writer
                    .write_event(Event::Text(BytesText::new(&n.to_string())))
                    .unwrap();
                writer
                    .write_event(Event::End(BytesEnd::new("vt:i4")))
                    .unwrap();
            }
            CustomPropertyValue::Float(f) => {
                writer
                    .write_event(Event::Start(BytesStart::new("vt:r8")))
                    .unwrap();
                writer
                    .write_event(Event::Text(BytesText::new(&f.to_string())))
                    .unwrap();
                writer
                    .write_event(Event::End(BytesEnd::new("vt:r8")))
                    .unwrap();
            }
            CustomPropertyValue::Bool(b) => {
                writer
                    .write_event(Event::Start(BytesStart::new("vt:bool")))
                    .unwrap();
                writer
                    .write_event(Event::Text(BytesText::new(if *b {
                        "true"
                    } else {
                        "false"
                    })))
                    .unwrap();
                writer
                    .write_event(Event::End(BytesEnd::new("vt:bool")))
                    .unwrap();
            }
            CustomPropertyValue::DateTime(dt) => {
                writer
                    .write_event(Event::Start(BytesStart::new("vt:filetime")))
                    .unwrap();
                writer.write_event(Event::Text(BytesText::new(dt))).unwrap();
                writer
                    .write_event(Event::End(BytesEnd::new("vt:filetime")))
                    .unwrap();
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new("property")))
            .unwrap();
    }

    writer
        .write_event(Event::End(BytesEnd::new("Properties")))
        .unwrap();

    String::from_utf8(writer.into_inner()).unwrap()
}

/// A `property` element whose closing tag has not been reached yet.
struct PendingProperty {
    fmtid: Option<String>,
    pid: Option<u32>,
    name: Option<String>,
    /// Element depth of the `property` element itself.
    depth: usize,
    /// Number of direct child elements seen; only the first is the value.
    value_children: usize,
    value_tag: Option<Vec<u8>>,
    value_text: String,
}

impl PendingProperty {
    fn from_start(e: &BytesStart<'_>, depth: usize) -> Result<Self, String> {
        let mut pending = Self {
            fmtid: None,
            pid: None,
            name: None,
            depth,
            value_children: 0,
            value_tag: None,
            value_text: String::new(),
        };
        for attr in e.attributes().flatten() {
            let slot = match attr.key.as_ref() {
                b"fmtid" => &mut pending.fmtid,
                b"name" => &mut pending.name,
                b"pid" => {
                    let raw = attr.unescape_value().map_err(xml_error)?;
                    pending.pid = raw.trim().parse().ok();
                    continue;
                }
                _ => continue,
            };
            *slot = Some(attr.unescape_value().map_err(xml_error)?.into_owned());
        }
        Ok(pending)
    }

    fn start_value(&mut self, local_name: &[u8]) {
        self.value_children += 1;
        if self.value_children == 1 {
            self.value_tag = Some(local_name.to_vec());
        }
    }

    fn finish(self) -> Option<CustomProperty> {
        let text = self.value_text;
        let value = match self.value_tag.as_deref()? {
            b"lpwstr" => CustomPropertyValue::String(text),
            b"i4" => CustomPropertyValue::Int(text.trim().parse().ok()?),
            b"r8" => CustomPropertyValue::Float(text.trim().parse().ok()?),
            b"bool" => {
                let t = text.trim();
                CustomPropertyValue::Bool(t == "true" || t == "1")
            }
            b"filetime" => CustomPropertyValue::DateTime(text.trim().to_string()),
            _ => return None,
        };
        Some(CustomProperty {
            fmtid: self.fmtid?,
            pid: self.pid?,
            name: self.name?,
            value,
        })
    }
}

/// Deserialize `CustomProperties` from an XML string.
///
/// Each `property` element is finalized at its closing tag, so empty,
/// whitespace-only, and CDATA string values are preserved. Attribute values
/// are entity-decoded. Value elements are matched by local name regardless of
/// prefix; properties with unsupported value types (e.g. `vt:vector`) are
/// skipped.
pub fn deserialize_custom_properties(xml: &str) -> Result<CustomProperties, String> {
    let mut reader = Reader::from_str(xml);

    let mut props = CustomProperties::default();
    let mut depth = 0usize;
    let mut pending: Option<PendingProperty> = None;

    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match &event {
            Event::Start(e) => {
                depth += 1;
                match pending.as_mut() {
                    None if e.local_name().as_ref() == b"property" => {
                        pending = Some(PendingProperty::from_start(e, depth)?);
                    }
                    Some(p) if depth == p.depth + 1 => p.start_value(e.local_name().as_ref()),
                    _ => {}
                }
            }
            Event::Empty(e) => {
                if let Some(p) = pending.as_mut() {
                    if depth == p.depth {
                        p.start_value(e.local_name().as_ref());
                    }
                }
            }
            Event::Text(_) | Event::CData(_) => {
                if let Some(p) = pending.as_mut() {
                    if depth == p.depth + 1 && p.value_children == 1 {
                        append_text(&mut p.value_text, &event)?;
                    }
                }
            }
            Event::End(_) => {
                if pending.as_ref().is_some_and(|p| p.depth == depth) {
                    if let Some(prop) = pending.take().and_then(PendingProperty::finish) {
                        props.properties.push(prop);
                    }
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

fn default_extended_properties_ns() -> String {
    namespaces::EXTENDED_PROPERTIES.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_core_properties_roundtrip() {
        let props = CoreProperties {
            title: Some("Test Title".to_string()),
            subject: Some("Test Subject".to_string()),
            creator: Some("Test Author".to_string()),
            keywords: Some("key1, key2".to_string()),
            description: Some("A description".to_string()),
            last_modified_by: Some("Editor".to_string()),
            revision: Some("3".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            modified: Some("2024-06-15T12:30:00Z".to_string()),
            category: Some("Reports".to_string()),
            content_status: Some("Draft".to_string()),
        };

        let xml = serialize_core_properties(&props);
        let parsed = deserialize_core_properties(&xml).unwrap();
        assert_eq!(props, parsed);
    }

    #[test]
    fn test_core_properties_empty_fields() {
        let props = CoreProperties::default();
        let xml = serialize_core_properties(&props);
        let parsed = deserialize_core_properties(&xml).unwrap();
        assert_eq!(props, parsed);
    }

    #[test]
    fn test_core_properties_partial_fields() {
        let props = CoreProperties {
            title: Some("Only Title".to_string()),
            creator: Some("Only Author".to_string()),
            ..Default::default()
        };

        let xml = serialize_core_properties(&props);
        let parsed = deserialize_core_properties(&xml).unwrap();
        assert_eq!(props, parsed);
    }

    #[test]
    fn test_core_properties_serialized_format() {
        let props = CoreProperties {
            title: Some("My Title".to_string()),
            creator: Some("Author Name".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            ..Default::default()
        };

        let xml = serialize_core_properties(&props);
        assert!(xml.contains("<cp:coreProperties"));
        assert!(xml.contains("xmlns:cp="));
        assert!(xml.contains("xmlns:dc="));
        assert!(xml.contains("xmlns:dcterms="));
        assert!(xml.contains("<dc:title>My Title</dc:title>"));
        assert!(xml.contains("<dc:creator>Author Name</dc:creator>"));
        assert!(xml.contains("xsi:type=\"dcterms:W3CDTF\""));
        assert!(xml.contains("<dcterms:created"));
        assert!(xml.contains("2024-01-01T00:00:00Z</dcterms:created>"));
    }

    #[test]
    fn test_parse_real_excel_core_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Budget Report</dc:title>
  <dc:subject>Finance</dc:subject>
  <dc:creator>John Doe</dc:creator>
  <cp:keywords>budget, 2024</cp:keywords>
  <dc:description>Annual budget report</dc:description>
  <cp:lastModifiedBy>Jane Smith</cp:lastModifiedBy>
  <cp:revision>5</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T08:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-20T16:45:00Z</dcterms:modified>
  <cp:category>Financial</cp:category>
  <cp:contentStatus>Final</cp:contentStatus>
</cp:coreProperties>"#;

        let props = deserialize_core_properties(xml).unwrap();
        assert_eq!(props.title.as_deref(), Some("Budget Report"));
        assert_eq!(props.subject.as_deref(), Some("Finance"));
        assert_eq!(props.creator.as_deref(), Some("John Doe"));
        assert_eq!(props.keywords.as_deref(), Some("budget, 2024"));
        assert_eq!(props.description.as_deref(), Some("Annual budget report"));
        assert_eq!(props.last_modified_by.as_deref(), Some("Jane Smith"));
        assert_eq!(props.revision.as_deref(), Some("5"));
        assert_eq!(props.created.as_deref(), Some("2024-01-15T08:00:00Z"));
        assert_eq!(props.modified.as_deref(), Some("2024-06-20T16:45:00Z"));
        assert_eq!(props.category.as_deref(), Some("Financial"));
        assert_eq!(props.content_status.as_deref(), Some("Final"));
    }

    #[test]
    fn test_extended_properties_serde_roundtrip() {
        let props = ExtendedProperties {
            xmlns: namespaces::EXTENDED_PROPERTIES.to_string(),
            xmlns_vt: Some(namespaces::VT.to_string()),
            application: Some("SheetKit".to_string()),
            doc_security: Some(0),
            scale_crop: Some(false),
            company: Some("Acme Corp".to_string()),
            links_up_to_date: Some(false),
            shared_doc: Some(false),
            hyperlinks_changed: Some(false),
            app_version: Some("1.0.0".to_string()),
            template: None,
            manager: Some("Boss".to_string()),
        };

        let xml = quick_xml::se::to_string(&props).unwrap();
        let parsed: ExtendedProperties = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(props, parsed);
    }

    #[test]
    fn test_extended_properties_with_defaults() {
        let props = ExtendedProperties::with_defaults();
        assert_eq!(props.xmlns, namespaces::EXTENDED_PROPERTIES);
        assert_eq!(props.xmlns_vt.as_deref(), Some(namespaces::VT));
        assert!(props.application.is_none());
    }

    #[test]
    fn test_extended_properties_skip_none_fields() {
        let props = ExtendedProperties {
            xmlns: namespaces::EXTENDED_PROPERTIES.to_string(),
            xmlns_vt: None,
            application: Some("Test".to_string()),
            doc_security: None,
            scale_crop: None,
            company: None,
            links_up_to_date: None,
            shared_doc: None,
            hyperlinks_changed: None,
            app_version: None,
            template: None,
            manager: None,
        };

        let xml = quick_xml::se::to_string(&props).unwrap();
        assert!(xml.contains("<Application>Test</Application>"));
        assert!(!xml.contains("DocSecurity"));
        assert!(!xml.contains("Company"));
    }

    #[test]
    fn test_custom_properties_roundtrip() {
        let props = CustomProperties {
            properties: vec![
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 2,
                    name: "Project".to_string(),
                    value: CustomPropertyValue::String("SheetKit".to_string()),
                },
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 3,
                    name: "Version".to_string(),
                    value: CustomPropertyValue::Int(42),
                },
            ],
        };

        let xml = serialize_custom_properties(&props);
        let parsed = deserialize_custom_properties(&xml).unwrap();
        assert_eq!(props, parsed);
    }

    #[test]
    fn test_custom_properties_all_value_types() {
        let props = CustomProperties {
            properties: vec![
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 2,
                    name: "StringProp".to_string(),
                    value: CustomPropertyValue::String("hello".to_string()),
                },
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 3,
                    name: "IntProp".to_string(),
                    value: CustomPropertyValue::Int(-7),
                },
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 4,
                    name: "FloatProp".to_string(),
                    value: CustomPropertyValue::Float(3.15),
                },
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 5,
                    name: "BoolProp".to_string(),
                    value: CustomPropertyValue::Bool(true),
                },
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 6,
                    name: "DateProp".to_string(),
                    value: CustomPropertyValue::DateTime("2024-01-01T00:00:00Z".to_string()),
                },
            ],
        };

        let xml = serialize_custom_properties(&props);
        let parsed = deserialize_custom_properties(&xml).unwrap();
        assert_eq!(props.properties.len(), parsed.properties.len());
        for (orig, p) in props.properties.iter().zip(parsed.properties.iter()) {
            assert_eq!(orig.name, p.name);
            assert_eq!(orig.value, p.value);
        }
    }

    #[test]
    fn test_custom_properties_empty() {
        let props = CustomProperties::default();
        let xml = serialize_custom_properties(&props);
        let parsed = deserialize_custom_properties(&xml).unwrap();
        assert!(parsed.properties.is_empty());
    }

    #[test]
    fn test_custom_properties_serialized_format() {
        let props = CustomProperties {
            properties: vec![CustomProperty {
                fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                pid: 2,
                name: "MyProp".to_string(),
                value: CustomPropertyValue::String("MyValue".to_string()),
            }],
        };

        let xml = serialize_custom_properties(&props);
        assert!(xml.contains("<Properties"));
        assert!(xml.contains("xmlns:vt="));
        assert!(xml.contains("<property"));
        assert!(xml.contains("fmtid="));
        assert!(xml.contains("pid=\"2\""));
        assert!(xml.contains("name=\"MyProp\""));
        assert!(xml.contains("<vt:lpwstr>MyValue</vt:lpwstr>"));
    }

    #[test]
    fn test_custom_properties_bool_false() {
        let props = CustomProperties {
            properties: vec![CustomProperty {
                fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                pid: 2,
                name: "Flag".to_string(),
                value: CustomPropertyValue::Bool(false),
            }],
        };

        let xml = serialize_custom_properties(&props);
        let parsed = deserialize_custom_properties(&xml).unwrap();
        assert_eq!(parsed.properties[0].value, CustomPropertyValue::Bool(false));
    }

    fn string_prop(pid: u32, name: &str, value: &str) -> CustomProperty {
        CustomProperty {
            fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
            pid,
            name: name.to_string(),
            value: CustomPropertyValue::String(value.to_string()),
        }
    }

    #[test]
    fn test_custom_properties_special_chars_and_empty_values_multi_cycle() {
        let props = CustomProperties {
            properties: vec![
                string_prop(2, "A&B", "x&y"),
                string_prop(3, "<lt>", "<v>"),
                string_prop(4, "q\"uote'apos", "\"'"),
                string_prop(5, "Empty", ""),
                string_prop(6, "Spaces", "  "),
                string_prop(7, "Padded", " a "),
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 8,
                    name: "Count&Co".to_string(),
                    value: CustomPropertyValue::Int(3),
                },
            ],
        };

        let mut current = props.clone();
        for _ in 0..3 {
            let xml = serialize_custom_properties(&current);
            current = deserialize_custom_properties(&xml).unwrap();
            assert_eq!(current, props);
        }
    }

    #[test]
    fn test_custom_properties_cdata_and_pretty_printed_values() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name='Cd&amp;ata'>
    <vt:lpwstr><![CDATA[a<b>&amp;]]></vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Mixed"><vt:lpwstr>x &amp; <![CDATA[<y>]]></vt:lpwstr></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="SelfClosed"><vt:lpwstr/></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Num">
    <vt:i4> 7 </vt:i4>
  </property>
</Properties>"#;

        let parsed = deserialize_custom_properties(xml).unwrap();
        assert_eq!(
            parsed.properties,
            vec![
                string_prop(2, "Cd&ata", "a<b>&amp;"),
                string_prop(3, "Mixed", "x & <y>"),
                string_prop(4, "SelfClosed", ""),
                CustomProperty {
                    fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                    pid: 5,
                    name: "Num".to_string(),
                    value: CustomPropertyValue::Int(7),
                },
            ]
        );

        let again = deserialize_custom_properties(&serialize_custom_properties(&parsed)).unwrap();
        assert_eq!(again, parsed);
    }

    #[test]
    fn test_custom_properties_vector_value_does_not_leak_inner_items() {
        let xml = r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Vec"><vt:vector size="1" baseType="lpwstr"><vt:lpwstr>inner</vt:lpwstr></vt:vector></property>
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="After"><vt:lpwstr>ok</vt:lpwstr></property>
</Properties>"#;

        let parsed = deserialize_custom_properties(xml).unwrap();
        assert_eq!(parsed.properties, vec![string_prop(3, "After", "ok")]);
    }

    #[test]
    fn test_core_properties_special_chars_and_empty_values_multi_cycle() {
        let props = CoreProperties {
            title: Some("A&B <c> \"d\" 'e'".to_string()),
            subject: Some(String::new()),
            creator: Some("  ".to_string()),
            keywords: Some(" k1, k2 ".to_string()),
            ..Default::default()
        };

        let mut current = props.clone();
        for _ in 0..3 {
            let xml = serialize_core_properties(&current);
            current = deserialize_core_properties(&xml).unwrap();
            assert_eq!(current, props);
        }
    }

    #[test]
    fn test_core_properties_cdata_and_alternate_prefixes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<core:coreProperties xmlns:core="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:d="http://purl.org/dc/elements/1.1/">
  <d:title><![CDATA[T<1>&]]></d:title>
  <d:creator/>
  <core:keywords>a &amp; b</core:keywords>
</core:coreProperties>"#;

        let props = deserialize_core_properties(xml).unwrap();
        assert_eq!(props.title.as_deref(), Some("T<1>&"));
        assert_eq!(props.creator.as_deref(), Some(""));
        assert_eq!(props.keywords.as_deref(), Some("a & b"));
        assert_eq!(props.subject, None);
    }

    #[test]
    fn test_extended_properties_special_chars_and_empty_values() {
        let mut props = ExtendedProperties::with_defaults();
        props.company = Some("A&B <c> \"d\" 'e'".to_string());
        props.manager = Some(String::new());

        let mut current = props.clone();
        for _ in 0..3 {
            let xml = quick_xml::se::to_string(&current).unwrap();
            current = quick_xml::de::from_str(&xml).unwrap();
            assert_eq!(current, props);
        }
    }
}
