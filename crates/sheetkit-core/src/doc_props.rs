//! User-facing document properties types and conversion helpers.
//!
//! These types provide an ergonomic API for working with the three kinds of
//! OOXML document properties (core, extended/application, and custom).

use sheetkit_xml::doc_props::{
    CoreProperties, CustomProperty, CustomPropertyValue as XmlCustomPropertyValue,
    ExtendedProperties, CUSTOM_PROPERTY_FMTID,
};
use sheetkit_xml::namespaces;
use sheetkit_xml::workbook::{CalcPr, WorkbookPr};

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

/// High-level workbook properties (maps to the `WorkbookPr` XML element).
///
/// All fields are optional; `None` means the attribute is omitted from the XML.
#[derive(Debug, Clone, Default)]
pub struct WorkbookSettings {
    /// Use 1904 date system instead of 1900.
    pub date1904: Option<bool>,
    /// Filter privacy setting.
    pub filter_privacy: Option<bool>,
    /// Default theme version.
    pub default_theme_version: Option<u32>,
    /// Show objects mode (e.g. "all", "placeholders", "none").
    pub show_objects: Option<String>,
    /// Code name for VBA.
    pub code_name: Option<String>,
    /// Check compatibility on save.
    pub check_compatibility: Option<bool>,
    /// Auto compress pictures.
    pub auto_compress_pictures: Option<bool>,
    /// Backup file setting.
    pub backup_file: Option<bool>,
    /// Save external link values.
    pub save_external_link_values: Option<bool>,
    /// Update links mode.
    pub update_links: Option<String>,
    /// Hide pivot field list.
    pub hide_pivot_field_list: Option<bool>,
    /// Show pivot chart filter.
    pub show_pivot_chart_filter: Option<bool>,
    /// Allow refresh query.
    pub allow_refresh_query: Option<bool>,
    /// Publish items.
    pub publish_items: Option<bool>,
    /// Show border on unselected tables.
    pub show_border_unselected_tables: Option<bool>,
    /// Prompted solutions.
    pub prompted_solutions: Option<bool>,
    /// Show ink annotation.
    pub show_ink_annotation: Option<bool>,
}

impl From<&WorkbookPr> for WorkbookSettings {
    fn from(pr: &WorkbookPr) -> Self {
        Self {
            date1904: pr.date1904,
            filter_privacy: pr.filter_privacy,
            default_theme_version: pr.default_theme_version,
            show_objects: pr.show_objects.clone(),
            code_name: pr.code_name.clone(),
            check_compatibility: pr.check_compatibility,
            auto_compress_pictures: pr.auto_compress_pictures,
            backup_file: pr.backup_file,
            save_external_link_values: pr.save_external_link_values,
            update_links: pr.update_links.clone(),
            hide_pivot_field_list: pr.hide_pivot_field_list,
            show_pivot_chart_filter: pr.show_pivot_chart_filter,
            allow_refresh_query: pr.allow_refresh_query,
            publish_items: pr.publish_items,
            show_border_unselected_tables: pr.show_border_unselected_tables,
            prompted_solutions: pr.prompted_solutions,
            show_ink_annotation: pr.show_ink_annotation,
        }
    }
}

impl WorkbookSettings {
    /// Convert to the XML-level `WorkbookPr` struct.
    pub fn to_workbook_pr(&self) -> WorkbookPr {
        WorkbookPr {
            date1904: self.date1904,
            filter_privacy: self.filter_privacy,
            default_theme_version: self.default_theme_version,
            show_objects: self.show_objects.clone(),
            code_name: self.code_name.clone(),
            check_compatibility: self.check_compatibility,
            auto_compress_pictures: self.auto_compress_pictures,
            backup_file: self.backup_file,
            save_external_link_values: self.save_external_link_values,
            update_links: self.update_links.clone(),
            hide_pivot_field_list: self.hide_pivot_field_list,
            show_pivot_chart_filter: self.show_pivot_chart_filter,
            allow_refresh_query: self.allow_refresh_query,
            publish_items: self.publish_items,
            show_border_unselected_tables: self.show_border_unselected_tables,
            prompted_solutions: self.prompted_solutions,
            show_ink_annotation: self.show_ink_annotation,
        }
    }
}

/// High-level calculation properties (maps to the `CalcPr` XML element).
///
/// All fields are optional; `None` means the attribute is omitted from the XML.
#[derive(Debug, Clone, Default)]
pub struct CalcSettings {
    /// Calculation engine ID.
    pub calc_id: Option<u32>,
    /// Calculation mode: "auto", "manual", or "autoNoTable".
    pub calc_mode: Option<String>,
    /// Full calculation on load.
    pub full_calc_on_load: Option<bool>,
    /// Reference mode: "A1" or "R1C1".
    pub ref_mode: Option<String>,
    /// Enable iterative calculation.
    pub iterate: Option<bool>,
    /// Maximum iterations for iterative calculation.
    pub iterate_count: Option<u32>,
    /// Delta threshold for iterative calculation convergence.
    pub iterate_delta: Option<f64>,
    /// Full precision for calculations.
    pub full_precision: Option<bool>,
    /// Whether calculation was completed before save.
    pub calc_completed: Option<bool>,
    /// Calculate on save.
    pub calc_on_save: Option<bool>,
    /// Enable concurrent calculation.
    pub concurrent_calc: Option<bool>,
    /// Manual concurrent calculation thread count.
    pub concurrent_manual_count: Option<u32>,
    /// Force full calculation.
    pub force_full_calc: Option<bool>,
}

impl From<&CalcPr> for CalcSettings {
    fn from(pr: &CalcPr) -> Self {
        Self {
            calc_id: pr.calc_id,
            calc_mode: pr.calc_mode.clone(),
            full_calc_on_load: pr.full_calc_on_load,
            ref_mode: pr.ref_mode.clone(),
            iterate: pr.iterate,
            iterate_count: pr.iterate_count,
            iterate_delta: pr.iterate_delta,
            full_precision: pr.full_precision,
            calc_completed: pr.calc_completed,
            calc_on_save: pr.calc_on_save,
            concurrent_calc: pr.concurrent_calc,
            concurrent_manual_count: pr.concurrent_manual_count,
            force_full_calc: pr.force_full_calc,
        }
    }
}

impl CalcSettings {
    /// Convert to the XML-level `CalcPr` struct.
    pub fn to_calc_pr(&self) -> CalcPr {
        CalcPr {
            calc_id: self.calc_id,
            calc_mode: self.calc_mode.clone(),
            full_calc_on_load: self.full_calc_on_load,
            ref_mode: self.ref_mode.clone(),
            iterate: self.iterate,
            iterate_count: self.iterate_count,
            iterate_delta: self.iterate_delta,
            full_precision: self.full_precision,
            calc_completed: self.calc_completed,
            calc_on_save: self.calc_on_save,
            concurrent_calc: self.concurrent_calc,
            concurrent_manual_count: self.concurrent_manual_count,
            force_full_calc: self.force_full_calc,
        }
    }
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

    // ---- WorkbookSettings tests ----

    #[test]
    fn test_workbook_settings_default() {
        let settings = WorkbookSettings::default();
        assert!(settings.date1904.is_none());
        assert!(settings.filter_privacy.is_none());
        assert!(settings.default_theme_version.is_none());
        assert!(settings.show_objects.is_none());
        assert!(settings.code_name.is_none());
        assert!(settings.check_compatibility.is_none());
        assert!(settings.auto_compress_pictures.is_none());
        assert!(settings.backup_file.is_none());
        assert!(settings.save_external_link_values.is_none());
        assert!(settings.update_links.is_none());
        assert!(settings.hide_pivot_field_list.is_none());
        assert!(settings.show_pivot_chart_filter.is_none());
        assert!(settings.allow_refresh_query.is_none());
        assert!(settings.publish_items.is_none());
        assert!(settings.show_border_unselected_tables.is_none());
        assert!(settings.prompted_solutions.is_none());
        assert!(settings.show_ink_annotation.is_none());
    }

    #[test]
    fn test_workbook_settings_to_xml_roundtrip() {
        let settings = WorkbookSettings {
            date1904: Some(false),
            filter_privacy: Some(true),
            default_theme_version: Some(166925),
            show_objects: Some("all".to_string()),
            code_name: Some("ThisWorkbook".to_string()),
            check_compatibility: Some(true),
            auto_compress_pictures: Some(false),
            backup_file: Some(true),
            save_external_link_values: Some(true),
            update_links: Some("always".to_string()),
            hide_pivot_field_list: Some(false),
            show_pivot_chart_filter: Some(true),
            allow_refresh_query: Some(true),
            publish_items: Some(false),
            show_border_unselected_tables: Some(true),
            prompted_solutions: Some(false),
            show_ink_annotation: Some(true),
        };
        let pr = settings.to_workbook_pr();
        let back = WorkbookSettings::from(&pr);

        assert_eq!(back.date1904, Some(false));
        assert_eq!(back.filter_privacy, Some(true));
        assert_eq!(back.default_theme_version, Some(166925));
        assert_eq!(back.show_objects.as_deref(), Some("all"));
        assert_eq!(back.code_name.as_deref(), Some("ThisWorkbook"));
        assert_eq!(back.check_compatibility, Some(true));
        assert_eq!(back.auto_compress_pictures, Some(false));
        assert_eq!(back.backup_file, Some(true));
        assert_eq!(back.save_external_link_values, Some(true));
        assert_eq!(back.update_links.as_deref(), Some("always"));
        assert_eq!(back.hide_pivot_field_list, Some(false));
        assert_eq!(back.show_pivot_chart_filter, Some(true));
        assert_eq!(back.allow_refresh_query, Some(true));
        assert_eq!(back.publish_items, Some(false));
        assert_eq!(back.show_border_unselected_tables, Some(true));
        assert_eq!(back.prompted_solutions, Some(false));
        assert_eq!(back.show_ink_annotation, Some(true));
    }

    #[test]
    fn test_workbook_settings_date1904() {
        let settings = WorkbookSettings {
            date1904: Some(true),
            ..Default::default()
        };
        let pr = settings.to_workbook_pr();
        assert_eq!(pr.date1904, Some(true));
        // All other fields should be None
        assert!(pr.filter_privacy.is_none());
        assert!(pr.default_theme_version.is_none());
        assert!(pr.code_name.is_none());

        let back = WorkbookSettings::from(&pr);
        assert_eq!(back.date1904, Some(true));
        assert!(back.filter_privacy.is_none());
    }

    // ---- CalcSettings tests ----

    #[test]
    fn test_calc_settings_default() {
        let settings = CalcSettings::default();
        assert!(settings.calc_id.is_none());
        assert!(settings.calc_mode.is_none());
        assert!(settings.full_calc_on_load.is_none());
        assert!(settings.ref_mode.is_none());
        assert!(settings.iterate.is_none());
        assert!(settings.iterate_count.is_none());
        assert!(settings.iterate_delta.is_none());
        assert!(settings.full_precision.is_none());
        assert!(settings.calc_completed.is_none());
        assert!(settings.calc_on_save.is_none());
        assert!(settings.concurrent_calc.is_none());
        assert!(settings.concurrent_manual_count.is_none());
        assert!(settings.force_full_calc.is_none());
    }

    #[test]
    fn test_calc_settings_to_xml_roundtrip() {
        let settings = CalcSettings {
            calc_id: Some(191029),
            calc_mode: Some("auto".to_string()),
            full_calc_on_load: Some(true),
            ref_mode: Some("A1".to_string()),
            iterate: Some(true),
            iterate_count: Some(100),
            iterate_delta: Some(0.001),
            full_precision: Some(true),
            calc_completed: Some(true),
            calc_on_save: Some(true),
            concurrent_calc: Some(true),
            concurrent_manual_count: Some(4),
            force_full_calc: Some(false),
        };
        let pr = settings.to_calc_pr();
        let back = CalcSettings::from(&pr);

        assert_eq!(back.calc_id, Some(191029));
        assert_eq!(back.calc_mode.as_deref(), Some("auto"));
        assert_eq!(back.full_calc_on_load, Some(true));
        assert_eq!(back.ref_mode.as_deref(), Some("A1"));
        assert_eq!(back.iterate, Some(true));
        assert_eq!(back.iterate_count, Some(100));
        assert_eq!(back.iterate_delta, Some(0.001));
        assert_eq!(back.full_precision, Some(true));
        assert_eq!(back.calc_completed, Some(true));
        assert_eq!(back.calc_on_save, Some(true));
        assert_eq!(back.concurrent_calc, Some(true));
        assert_eq!(back.concurrent_manual_count, Some(4));
        assert_eq!(back.force_full_calc, Some(false));
    }

    #[test]
    fn test_calc_settings_manual_mode() {
        let settings = CalcSettings {
            calc_mode: Some("manual".to_string()),
            calc_on_save: Some(false),
            ..Default::default()
        };
        let pr = settings.to_calc_pr();
        assert_eq!(pr.calc_mode.as_deref(), Some("manual"));
        assert_eq!(pr.calc_on_save, Some(false));
        // All other fields should be None
        assert!(pr.calc_id.is_none());
        assert!(pr.iterate.is_none());

        let back = CalcSettings::from(&pr);
        assert_eq!(back.calc_mode.as_deref(), Some("manual"));
        assert_eq!(back.calc_on_save, Some(false));
        assert!(back.calc_id.is_none());
    }

    #[test]
    fn test_calc_settings_iterative() {
        let settings = CalcSettings {
            iterate: Some(true),
            iterate_count: Some(200),
            iterate_delta: Some(0.0001),
            ..Default::default()
        };
        let pr = settings.to_calc_pr();
        assert_eq!(pr.iterate, Some(true));
        assert_eq!(pr.iterate_count, Some(200));
        assert_eq!(pr.iterate_delta, Some(0.0001));
        // Other fields None
        assert!(pr.calc_mode.is_none());
        assert!(pr.ref_mode.is_none());

        let back = CalcSettings::from(&pr);
        assert_eq!(back.iterate, Some(true));
        assert_eq!(back.iterate_count, Some(200));
        assert_eq!(back.iterate_delta, Some(0.0001));
    }
}
