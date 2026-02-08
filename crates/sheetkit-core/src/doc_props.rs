//! User-facing document properties types and conversion helpers.
//!
//! These types provide an ergonomic API for working with the three kinds of
//! OOXML document properties (core, extended/application, and custom).

use sheetkit_xml::doc_props::{
    CoreProperties, CustomProperty, CustomPropertyValue as XmlCustomPropertyValue,
    ExtendedProperties, CUSTOM_PROPERTY_FMTID,
};
use sheetkit_xml::namespaces;

/// User-facing document core properties.
#[derive(Debug, Clone, Default)]
pub struct DocProperties {
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

impl From<&CoreProperties> for DocProperties {
    fn from(props: &CoreProperties) -> Self {
        Self {
            title: props.title.clone(),
            subject: props.subject.clone(),
            creator: props.creator.clone(),
            keywords: props.keywords.clone(),
            description: props.description.clone(),
            last_modified_by: props.last_modified_by.clone(),
            revision: props.revision.clone(),
            created: props.created.clone(),
            modified: props.modified.clone(),
            category: props.category.clone(),
            content_status: props.content_status.clone(),
        }
    }
}

impl DocProperties {
    /// Convert to the XML-level `CoreProperties` struct.
    pub fn to_core_properties(&self) -> CoreProperties {
        CoreProperties {
            title: self.title.clone(),
            subject: self.subject.clone(),
            creator: self.creator.clone(),
            keywords: self.keywords.clone(),
            description: self.description.clone(),
            last_modified_by: self.last_modified_by.clone(),
            revision: self.revision.clone(),
            created: self.created.clone(),
            modified: self.modified.clone(),
            category: self.category.clone(),
            content_status: self.content_status.clone(),
        }
    }
}

/// User-facing application properties.
#[derive(Debug, Clone, Default)]
pub struct AppProperties {
    pub application: Option<String>,
    pub doc_security: Option<u32>,
    pub company: Option<String>,
    pub app_version: Option<String>,
    pub manager: Option<String>,
    pub template: Option<String>,
}

impl From<&ExtendedProperties> for AppProperties {
    fn from(props: &ExtendedProperties) -> Self {
        Self {
            application: props.application.clone(),
            doc_security: props.doc_security,
            company: props.company.clone(),
            app_version: props.app_version.clone(),
            manager: props.manager.clone(),
            template: props.template.clone(),
        }
    }
}

impl AppProperties {
    /// Convert to the XML-level `ExtendedProperties` struct.
    pub fn to_extended_properties(&self) -> ExtendedProperties {
        ExtendedProperties {
            xmlns: namespaces::EXTENDED_PROPERTIES.to_string(),
            xmlns_vt: Some(namespaces::VT.to_string()),
            application: self.application.clone(),
            doc_security: self.doc_security,
            scale_crop: None,
            company: self.company.clone(),
            links_up_to_date: None,
            shared_doc: None,
            hyperlinks_changed: None,
            app_version: self.app_version.clone(),
            template: self.template.clone(),
            manager: self.manager.clone(),
        }
    }
}

/// Value type for custom properties.
#[derive(Debug, Clone, PartialEq)]
pub enum CustomPropertyValue {
    String(String),
    Int(i32),
    Float(f64),
    Bool(bool),
    DateTime(String),
}

impl CustomPropertyValue {
    /// Convert to the XML-level `CustomPropertyValue`.
    pub(crate) fn to_xml(&self) -> XmlCustomPropertyValue {
        match self {
            Self::String(s) => XmlCustomPropertyValue::String(s.clone()),
            Self::Int(n) => XmlCustomPropertyValue::Int(*n),
            Self::Float(f) => XmlCustomPropertyValue::Float(*f),
            Self::Bool(b) => XmlCustomPropertyValue::Bool(*b),
            Self::DateTime(dt) => XmlCustomPropertyValue::DateTime(dt.clone()),
        }
    }

    /// Convert from the XML-level `CustomPropertyValue`.
    pub(crate) fn from_xml(val: &XmlCustomPropertyValue) -> Self {
        match val {
            XmlCustomPropertyValue::String(s) => Self::String(s.clone()),
            XmlCustomPropertyValue::Int(n) => Self::Int(*n),
            XmlCustomPropertyValue::Float(f) => Self::Float(*f),
            XmlCustomPropertyValue::Bool(b) => Self::Bool(*b),
            XmlCustomPropertyValue::DateTime(dt) => Self::DateTime(dt.clone()),
        }
    }
}

/// Find a custom property by name and return its value, or None.
pub(crate) fn find_custom_property(
    props: &sheetkit_xml::doc_props::CustomProperties,
    name: &str,
) -> Option<CustomPropertyValue> {
    props
        .properties
        .iter()
        .find(|p| p.name == name)
        .map(|p| CustomPropertyValue::from_xml(&p.value))
}

/// Set a custom property by name (insert or update).
/// Returns the next available pid.
pub(crate) fn set_custom_property(
    props: &mut sheetkit_xml::doc_props::CustomProperties,
    name: &str,
    value: CustomPropertyValue,
) {
    if let Some(existing) = props.properties.iter_mut().find(|p| p.name == name) {
        existing.value = value.to_xml();
        return;
    }

    let next_pid = props
        .properties
        .iter()
        .map(|p| p.pid)
        .max()
        .map(|m| m + 1)
        .unwrap_or(2);

    props.properties.push(CustomProperty {
        fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
        pid: next_pid,
        name: name.to_string(),
        value: value.to_xml(),
    });
}

/// Remove a custom property by name. Returns true if found and removed.
pub(crate) fn delete_custom_property(
    props: &mut sheetkit_xml::doc_props::CustomProperties,
    name: &str,
) -> bool {
    let before = props.properties.len();
    props.properties.retain(|p| p.name != name);
    props.properties.len() < before
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_doc_properties_from_core_properties() {
        let core = CoreProperties {
            title: Some("T".to_string()),
            creator: Some("C".to_string()),
            ..Default::default()
        };
        let doc = DocProperties::from(&core);
        assert_eq!(doc.title.as_deref(), Some("T"));
        assert_eq!(doc.creator.as_deref(), Some("C"));
        assert!(doc.subject.is_none());
    }

    #[test]
    fn test_doc_properties_to_core_properties() {
        let doc = DocProperties {
            title: Some("T".to_string()),
            subject: Some("S".to_string()),
            ..Default::default()
        };
        let core = doc.to_core_properties();
        assert_eq!(core.title.as_deref(), Some("T"));
        assert_eq!(core.subject.as_deref(), Some("S"));
        assert!(core.creator.is_none());
    }

    #[test]
    fn test_app_properties_from_extended_properties() {
        let ext = ExtendedProperties {
            xmlns: namespaces::EXTENDED_PROPERTIES.to_string(),
            xmlns_vt: None,
            application: Some("TestApp".to_string()),
            doc_security: Some(0),
            company: Some("Corp".to_string()),
            ..Default::default()
        };
        let app = AppProperties::from(&ext);
        assert_eq!(app.application.as_deref(), Some("TestApp"));
        assert_eq!(app.doc_security, Some(0));
        assert_eq!(app.company.as_deref(), Some("Corp"));
    }

    #[test]
    fn test_app_properties_to_extended_properties() {
        let app = AppProperties {
            application: Some("SheetKit".to_string()),
            company: Some("Acme".to_string()),
            ..Default::default()
        };
        let ext = app.to_extended_properties();
        assert_eq!(ext.xmlns, namespaces::EXTENDED_PROPERTIES);
        assert_eq!(ext.application.as_deref(), Some("SheetKit"));
        assert_eq!(ext.company.as_deref(), Some("Acme"));
    }

    #[test]
    fn test_custom_property_value_roundtrip() {
        let vals = vec![
            CustomPropertyValue::String("hello".to_string()),
            CustomPropertyValue::Int(42),
            CustomPropertyValue::Float(3.14),
            CustomPropertyValue::Bool(true),
            CustomPropertyValue::DateTime("2024-01-01T00:00:00Z".to_string()),
        ];
        for v in &vals {
            let xml = v.to_xml();
            let back = CustomPropertyValue::from_xml(&xml);
            assert_eq!(*v, back);
        }
    }

    #[test]
    fn test_set_and_find_custom_property() {
        let mut props = sheetkit_xml::doc_props::CustomProperties::default();
        set_custom_property(
            &mut props,
            "Project",
            CustomPropertyValue::String("SK".to_string()),
        );
        let found = find_custom_property(&props, "Project");
        assert_eq!(found, Some(CustomPropertyValue::String("SK".to_string())));
        assert_eq!(props.properties[0].pid, 2);
    }

    #[test]
    fn test_set_custom_property_update_existing() {
        let mut props = sheetkit_xml::doc_props::CustomProperties::default();
        set_custom_property(
            &mut props,
            "Key",
            CustomPropertyValue::String("old".to_string()),
        );
        set_custom_property(
            &mut props,
            "Key",
            CustomPropertyValue::String("new".to_string()),
        );
        assert_eq!(props.properties.len(), 1);
        assert_eq!(
            find_custom_property(&props, "Key"),
            Some(CustomPropertyValue::String("new".to_string()))
        );
    }

    #[test]
    fn test_delete_custom_property() {
        let mut props = sheetkit_xml::doc_props::CustomProperties::default();
        set_custom_property(&mut props, "Key", CustomPropertyValue::Int(1));
        assert!(delete_custom_property(&mut props, "Key"));
        assert!(!delete_custom_property(&mut props, "Key")); // already gone
        assert!(find_custom_property(&props, "Key").is_none());
    }

    #[test]
    fn test_custom_property_pid_auto_increment() {
        let mut props = sheetkit_xml::doc_props::CustomProperties::default();
        set_custom_property(&mut props, "A", CustomPropertyValue::Int(1));
        set_custom_property(&mut props, "B", CustomPropertyValue::Int(2));
        set_custom_property(&mut props, "C", CustomPropertyValue::Int(3));
        assert_eq!(props.properties[0].pid, 2);
        assert_eq!(props.properties[1].pid, 3);
        assert_eq!(props.properties[2].pid, 4);
    }
}
